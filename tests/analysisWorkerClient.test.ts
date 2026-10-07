import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import { AnalysisWorkerClient } from '../src/analysisWorkerClient';

// Real worker threads against throwaway worker scripts: the client's contract
// is that failures never strand a caller. Startup failures allow the ordinary
// fallback; a hung analysis is rejected and the worker restarts off-thread.

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
    it('cancels a running obsolete live snapshot and dispatches the latest one', async () => {
        const marker = path.join(tempDir, 'running');
        client = new AnalysisWorkerClient(workerScript(`
            const { parentPort } = require('worker_threads');
            const fs = require('fs');
            parentPort.on('message', request => {
                if (request.kind !== 'analyze') return;
                if (request.source === 'obsolete') {
                    fs.writeFileSync(${JSON.stringify(marker)}, 'running');
                    const deadline = Date.now() + 3000;
                    while (Atomics.load(request.cancellationSignal, 0) === 0 && Date.now() < deadline) {}
                    parentPort.postMessage({ kind: Atomics.load(request.cancellationSignal, 0) ? 'cancelled' : 'error',
                        requestId: request.requestId, docKey: request.docKey, message: 'obsolete work was not cancelled' });
                } else {
                    parentPort.postMessage({ kind: 'result', requestId: request.requestId, docKey: request.docKey,
                        diagnostics: [], suppressedDiagnostics: [], incrementalMode: 'full' });
                }
            });
        `));
        const obsolete = client.analyze({ latestOnly: true, docKey: 'doc', source: 'obsolete', moduleName: 'M' })
            .then(() => undefined, error => error as Error);
        const deadline = Date.now() + 3000;
        while (!fs.existsSync(marker) && Date.now() < deadline) { await new Promise(resolve => setTimeout(resolve, 10)); }
        expect(fs.existsSync(marker)).toBe(true);
        const queued = client.analyze({ latestOnly: true, docKey: 'doc', source: 'queued', moduleName: 'M' })
            .then(() => undefined, error => error as Error);
        const current = await client.analyze({ latestOnly: true, docKey: 'doc', source: 'current', moduleName: 'M' });
        expect((await obsolete)?.name).toBe('AnalysisSnapshotSuperseded');
        expect((await queued)?.name).toBe('AnalysisSnapshotSuperseded');
        expect(current.diagnostics).toEqual([]);
        expect(client.available).toBe(true);
    });

    it('retains an explicit analysis when a live request for the same document arrives', async () => {
        client = new AnalysisWorkerClient(workerScript(ECHO_WORKER.replace("parentPort.postMessage({", "setTimeout(() => parentPort.postMessage({")
            .replace("incrementalMode: 'full',\n        });", "incrementalMode: 'full',\n        }), 40);")));
        const explicit = client.analyze({ docKey: 'doc', source: 'explicit', moduleName: 'M' });
        const live = client.analyze({ latestOnly: true, docKey: 'doc', source: 'live', moduleName: 'M' });
        const results = await Promise.all([explicit, live]);
        expect(results.every(result => result.incrementalMode === 'full')).toBe(true);
    });

    it('retains a live analysis for another document', async () => {
        client = new AnalysisWorkerClient(workerScript(ECHO_WORKER));
        const first = client.analyze({ latestOnly: true, docKey: 'first', source: 'first', moduleName: 'M' });
        const second = client.analyze({ latestOnly: true, docKey: 'second', source: 'second', moduleName: 'M' });
        expect((await Promise.all([first, second])).every(result => result.incrementalMode === 'full')).toBe(true);
    });

    it('cancels a forgotten live document without disabling the worker', async () => {
        client = new AnalysisWorkerClient(workerScript(ECHO_WORKER));
        const forgotten = client.analyze({ latestOnly: true, docKey: 'first', source: 'first', moduleName: 'M' })
            .then(() => undefined, error => error as Error);
        client.forget('first');
        expect((await forgotten)?.name).toBe('AnalysisSnapshotSuperseded');
        expect((await client.analyze({ latestOnly: true, docKey: 'second', source: 'second', moduleName: 'M' })).diagnostics).toEqual([]);
        expect(client.available).toBe(true);
    });
    it('resolves through a responsive worker', async () => {
        client = new AnalysisWorkerClient(workerScript(ECHO_WORKER));

        const result = await client.analyze({ docKey: 'doc', source: 'Sub A()\nEnd Sub', moduleName: 'M' });

        expect(result).toEqual({ diagnostics: [], suppressedDiagnostics: [], incrementalMode: 'full' });
        expect(client.available).toBe(true);
    });

    it('times out a hung analysis and runs the next request in a fresh worker', async () => {
        const marker = path.join(tempDir, 'started');
        client = new AnalysisWorkerClient(workerScript(`
            const fs = require('fs');
            if (fs.existsSync(${JSON.stringify(marker)})) {
                ${ECHO_WORKER}
            } else {
                fs.writeFileSync(${JSON.stringify(marker)}, 'started');
                ${SILENT_WORKER}
            }
        `), undefined, 200);

        await expect(client.analyze({ docKey: 'doc', source: 'Sub A()\nEnd Sub', moduleName: 'M' }))
            .rejects.toThrow(/timed out after 200 ms/);
        expect(client.available).toBe(true);
        expect((await client.analyze({ docKey: 'doc', source: 'Sub B()\nEnd Sub', moduleName: 'M' })).diagnostics).toEqual([]);
    });

    it('fails fast when the worker bundle does not exist', async () => {
        client = new AnalysisWorkerClient(path.join(tempDir, 'missing.js'));

        await expect(client.analyze({ docKey: 'doc', source: '', moduleName: 'M' }))
            .rejects.toThrow(/unavailable/);
        expect(client.available).toBe(false);
    });
});
