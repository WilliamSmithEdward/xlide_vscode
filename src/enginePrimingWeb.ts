// The browser half: read the container through workspace.fs before the
// engine runs, and write back what it produced.
//
// Cost control matters here. A workbook is a single file that can run to tens
// of megabytes, and the engine is called for every tree expansion, every
// module open, every analysis pass. Re-reading it each time would stall the
// editor, so the workspace's own stat decides: bytes are fetched only when
// the file's stamp differs from what is already primed. That is also what
// keeps the engine's own parse cache honest, since it keys on the same stamp
// (see cachedPackage in src/vba/projectService.ts).

import * as vscode from 'vscode';
import type { EnginePriming } from './enginePriming';
import { workspaceUriFor } from './util/workspaceUris';
import {
    primeHostFile,
    primeHostFileAbsent,
    primedStamp,
    takeHostFileWrites,
} from './vba/hostPlatformWeb';

export const enginePriming: EnginePriming = {
    async prime(filePaths: readonly string[]): Promise<void> {
        for (const filePath of filePaths) {
            const uri = workspaceUriFor(filePath);

            let stat: vscode.FileStat;
            try {
                stat = await vscode.workspace.fs.stat(uri);
            } catch {
                // The engine's callers treat a missing container as an error
                // worth reporting, and priming it as absent lets them say so
                // rather than reporting that XLIDE failed to load it.
                primeHostFileAbsent(filePath);
                continue;
            }

            const primed = primedStamp(filePath);
            if (primed && primed.mtimeMs === stat.mtime && primed.size === stat.size) {
                continue;
            }

            primeHostFile(filePath, Buffer.from(await vscode.workspace.fs.readFile(uri)), stat.mtime);
        }
    },

    async flush(): Promise<void> {
        for (const [filePath, data] of takeHostFileWrites()) {
            const uri = workspaceUriFor(filePath);
            await vscode.workspace.fs.writeFile(uri, data);
            try {
                // What is primed is byte-for-byte what was just written, so
                // adopting the workspace's new stamp saves re-reading the
                // whole container on the very next call. Failing to re-stat
                // only costs that read.
                const stat = await vscode.workspace.fs.stat(uri);
                primeHostFile(filePath, data, stat.mtime);
            } catch {
                /* the next prime will fetch it */
            }
        }
    },

    discard(): void {
        takeHostFileWrites();
    },
};
