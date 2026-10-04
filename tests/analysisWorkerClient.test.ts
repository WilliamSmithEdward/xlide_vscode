import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import { AnalysisWorkerClient } from '../src/analysisWorkerClient';

// Real worker threads against throwaway worker scripts: the client's contract
// is that a worker which cannot answer - missing, crashing, or silently hung -
// never strands a caller, because callers fall back to in-host analysis the
// moment `available` goes false.

let tempDir: string;
let client: AnalysisWorkerClient | undefined;

beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-worker-client-'));
});

afterEach(() => {
    client?.dispose();
    client = undefined;
    fs.rmSync(tempDir, { recursive: true, force: true });
});

function workerScript(body: string): string {
    const file = path.join(tempDir, 'worker.js');
    fs.writeFileSync(file, body, 'utf8');
    return file;
}

const ECHO_WORKER = `
const { parentPort } = require('worker_threads');
parentPort.on('message', (message) => {
    if (message.kind === 'analyze') {
        parentPort.postMessage({
            kind: 'result',
            requestId: message.requestId,
            docKey: message.docKey,
            diagnostics: [],
            suppressedDiagnostics: [],
            incrementalMode: 'full',
        });
    }
});
`;

// Listens so the thread stays alive, but never answers.
const SILENT_WORKER = `
const { parentPort } = require('worker_threads');
parentPort.on('message', () => {});
`;

describe('AnalysisWorkerClient', () => {
    it('resolves through a responsive worker', async () => {
        client = new AnalysisWorkerClient(workerScript(ECHO_WORKER));

        const result = await client.analyze({ docKey: 'doc', source: 'Sub A()\nEnd Sub', moduleName: 'M' });

        expect(result).toEqual({ diagnostics: [], suppressedDiagnostics: [], incrementalMode: 'full' });
        expect(client.available).toBe(true);
    });

    it('times out a hung worker and degrades to the in-host path', async () => {
        client = new AnalysisWorkerClient(workerScript(SILENT_WORKER), undefined, 100);

        await expect(client.analyze({ docKey: 'doc', source: 'Sub A()\nEnd Sub', moduleName: 'M' }))
            .rejects.toThrow(/timed out after 100 ms/);
        expect(client.available).toBe(false);
    });

    it('fails fast when the worker bundle does not exist', async () => {
        client = new AnalysisWorkerClient(path.join(tempDir, 'missing.js'));

        await expect(client.analyze({ docKey: 'doc', source: '', moduleName: 'M' }))
            .rejects.toThrow(/unavailable/);
        expect(client.available).toBe(false);
    });
});

const SLOW_WORKER = ECHO_WORKER.replace("if (message.kind === 'analyze') {", `if (message.kind === 'analyze') {
    const end = Date.now() + 80;
    while (Date.now() < end) {}
`);

describe('analysis queue', () => {
    it('does not count time waiting behind healthy requests as a timeout', async () => {
        client = new AnalysisWorkerClient(workerScript(SLOW_WORKER), undefined, 500);
        await Promise.all(Array.from({ length: 10 }, (_, i) => client!.analyze({
            docKey: String(i), source: '', moduleName: 'M',
        })));
        expect(client.available).toBe(true);
    });

    it('replaces queued live snapshots while preserving project requests', async () => {
        client = new AnalysisWorkerClient(workerScript(SLOW_WORKER.replace('diagnostics: [],', 'diagnostics: [{ message: message.source }],')));
        const first = client.analyze({ docKey: 'doc', source: 'first', moduleName: 'M', latestOnly: true });
        const stale = client.analyze({ docKey: 'doc', source: 'stale', moduleName: 'M', latestOnly: true });
        const rejected = expect(stale).rejects.toMatchObject({ name: 'AnalysisSnapshotSuperseded' });
        const project = client.analyze({ docKey: 'doc', source: 'project', moduleName: 'M' });
        const latest = client.analyze({ docKey: 'doc', source: 'latest', moduleName: 'M', latestOnly: true });
        await rejected;
        const results = await Promise.all([first, project, latest]);
        expect(results.map(r => r.diagnostics[0].message)).toEqual(['first', 'project', 'latest']);
        expect(client.available).toBe(true);
    });

    it('rejects both active and queued requests when disposed', async () => {
        client = new AnalysisWorkerClient(workerScript(SILENT_WORKER));
        const active = expect(client.analyze({ docKey: 'one', source: '', moduleName: 'M' })).rejects.toThrow('disposed');
        const queued = expect(client.analyze({ docKey: 'two', source: '', moduleName: 'M' })).rejects.toThrow('disposed');
        client.dispose();
        await Promise.all([active, queued]);
    });
});
