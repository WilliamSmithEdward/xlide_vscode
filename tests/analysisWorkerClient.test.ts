import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import { type WorkerAnalyzeRequest, AnalysisWorkerClient } from '../src/analysisWorkerClient';

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
        const first = expect(client.analyze({ docKey: 'doc', source: 'first', moduleName: 'M', latestOnly: true })).rejects.toMatchObject({ name: 'AnalysisSnapshotSuperseded' });
        const stale = client.analyze({ docKey: 'doc', source: 'stale', moduleName: 'M', latestOnly: true });
        const rejected = expect(stale).rejects.toMatchObject({ name: 'AnalysisSnapshotSuperseded' });
        const project = client.analyze({ docKey: 'doc', source: 'project', moduleName: 'M' });
        const latest = client.analyze({ docKey: 'doc', source: 'latest', moduleName: 'M', latestOnly: true });
        await rejected;
        await first;
        const results = await Promise.all([project, latest]);
        expect(results.map(r => r.diagnostics[0].message)).toEqual(['project', 'latest']);
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


describe('queue lifecycle', () => {
    it('drops queued analyses when their document closes', async () => {
        client = new AnalysisWorkerClient(workerScript(SLOW_WORKER));
        const active = client.analyze({ docKey: 'other', source: '', moduleName: 'M' });
        const queued = client.analyze({ docKey: 'closed', source: '', moduleName: 'M', latestOnly: true });
        const outcome = queued.then(() => 'analyzed', () => 'cancelled');
        client.forget('closed');
        await active;
        expect(await outcome).toBe('cancelled');
        expect(client.available).toBe(true);
    });

    it('settles a queued request when its seed provider throws', async () => {
        client = new AnalysisWorkerClient(workerScript(SLOW_WORKER));
        const active = client.analyze({ docKey: 'other', source: '', moduleName: 'M' });
        client.ensureSeeded('book', 1, () => { throw new Error('seed unavailable'); });
        const queued = client.analyze({ docKey: 'bad', projectKey: 'book', generation: 1, source: '', moduleName: 'M' });
        const rejected = expect(queued).rejects.toThrow('seed unavailable');
        const after = client.analyze({ docKey: 'after', source: '', moduleName: 'M' });
        await Promise.all([active, rejected, after]);
        expect(client.available).toBe(true);
    });
});


describe('seed dispatch', () => {
    it('retains each queued project generation and its seed provider', async () => {
        const seeded = SLOW_WORKER.replace("parentPort.on('message', (message) => {", `let seed;
parentPort.on('message', (message) => {
    if (message.kind === 'seed') { seed = message.modules[0].source; }`).replace('diagnostics: [],', 'diagnostics: [{ message: seed }],');
        client = new AnalysisWorkerClient(workerScript(seeded));
        const active = client.analyze({ docKey: 'other', source: '', moduleName: 'M' });
        client.ensureSeeded('book', 1, () => [{ moduleName: 'M', source: 'old' }]);
        const old = client.analyze({ docKey: 'one', projectKey: 'book', generation: 1, source: '', moduleName: 'M' });
        client.ensureSeeded('book', 2, () => [{ moduleName: 'M', source: 'new' }]);
        const next = client.analyze({ docKey: 'two', projectKey: 'book', generation: 2, source: '', moduleName: 'M' });
        await active;
        expect((await old).diagnostics[0].message).toBe('old');
        expect((await next).diagnostics[0].message).toBe('new');
    });

    it('settles a failed reseed and continues dispatching', async () => {
        const needSeed = ECHO_WORKER.replace("if (message.kind === 'analyze') {", `if (message.kind === 'analyze') {
            if (message.projectKey) { parentPort.postMessage({kind:'needSeed',requestId:message.requestId,projectKey:message.projectKey}); return; }`);
        client = new AnalysisWorkerClient(workerScript(needSeed));
        let calls = 0;
        client.ensureSeeded('book', 1, () => {
            if (++calls === 2) { throw new Error('reseed failed'); }
            return [{ moduleName: 'M', source: '' }];
        });
        const bad = expect(client.analyze({ docKey: 'one', projectKey: 'book', generation: 1, source: '', moduleName: 'M' })).rejects.toThrow('reseed failed');
        const next = client.analyze({ docKey: 'two', source: '', moduleName: 'M' });
        await Promise.all([bad, next]);
        expect(client.available).toBe(true);
    });
});


describe('request serialization failure', () => {
    it('rejects a non-cloneable queued request without retaining a watchdog or blocking the next', async () => {
        client = new AnalysisWorkerClient(workerScript(SLOW_WORKER), undefined, 500);
        const active = client.analyze({ docKey: 'active', source: '', moduleName: 'M' });
        const bad = client.analyze({ docKey: 'bad', source: '', moduleName: 'M', host: () => undefined } as unknown as WorkerAnalyzeRequest);
        const rejected = expect(bad).rejects.toThrow();
        const next = client.analyze({ docKey: 'next', source: '', moduleName: 'M' });
        await Promise.all([active, rejected, next]);
        expect(client.available).toBe(true);
    });
});


describe('timeout restart', () => {
    it('restarts after a genuinely stuck job and continues queued analysis with a fresh seed', async () => {
        const restarting = ECHO_WORKER.replace("parentPort.on('message', (message) => {", `let seeded = false;
parentPort.on('message', (message) => {
    if (message.kind === 'seed') { seeded = true; }`).replace("if (message.kind === 'analyze') {", `if (message.kind === 'analyze') {
    if (message.source === 'hang') { while (true) {} }
    if (message.projectKey && !seeded) { parentPort.postMessage({kind:'error',requestId:message.requestId,message:'missing seed'}); return; }`);
        client = new AnalysisWorkerClient(workerScript(restarting), undefined, 500);
        client.ensureSeeded('book', 1, () => [{ moduleName: 'M', source: '' }]);
        const failed = expect(client.analyze({ docKey: 'hung', source: 'hang', moduleName: 'M', projectKey: 'book', generation: 1 }))
            .rejects.toMatchObject({ name: 'AnalysisWorkerTimeoutError' });
        const queued = client.analyze({ docKey: 'healthy', source: '', moduleName: 'M', projectKey: 'book', generation: 1 });
        await Promise.all([failed, queued]);
        // The terminated worker's late exit must not disable its replacement.
        await client.analyze({ docKey: 'after', source: '', moduleName: 'M' });
        expect(client.available).toBe(true);
    });
});
