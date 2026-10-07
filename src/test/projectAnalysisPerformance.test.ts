import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, closeAllEditors, workspaceRoot } from './support';

// Opt-in real-workbook probe. The source workbook is copied before the command
// runs; no Office application or writes to the original workbook are needed.
// XLIDE_ANALYSIS_PERF_WORKBOOK=... npm run test:integration -- --grep "Run analysis performance"
suite('Run analysis performance', () => {
    test('completes a cold and cached run while the extension host stays responsive', async function () {
        const original = process.env.XLIDE_ANALYSIS_PERF_WORKBOOK;
        if (!original) { this.skip(); return; }
        this.timeout(120_000);
        await activate();
        await closeAllEditors();
        const filePath = path.join(workspaceRoot(), 'AnalysisPerformance.xlsm');
        fs.copyFileSync(original, filePath);
        let previous = Date.now();
        let maximumGap = 0;
        const heartbeat = setInterval(() => {
            const now = Date.now();
            maximumGap = Math.max(maximumGap, now - previous);
            previous = now;
        }, 25);
        const times: number[] = [];
        try {
            for (let i = 0; i < 2; i++) {
                const start = Date.now();
                await vscode.commands.executeCommand('xlide.analyzeProject', { filePath });
                times.push(Date.now() - start);
                await new Promise(resolve => setTimeout(resolve, 100));
            }
            await vscode.commands.executeCommand('xlide.copyPerformanceSnapshot');
            const snapshot = await vscode.env.clipboard.readText();
            console.log(JSON.stringify({ analysisPerformance: { times, maximumGap } }));
            console.log(snapshot);
            assert.equal((snapshot.match(/analyzeProject\.total ok/g) ?? []).length, 2);
            assert.doesNotMatch(snapshot, /analyzeProject\.total (failed|canceled)/);
            // The cache should be visible in the real command, and a slow worker
            // must not trigger the synchronous fallback that freezes the host.
            assert.ok(times[1] < 2000, `cached run took ${times[1]} ms`);
            assert.ok(maximumGap < 2000, `extension host stalled for ${maximumGap} ms`);
        } finally {
            clearInterval(heartbeat);
            await closeAllEditors();
        }
    });
});
