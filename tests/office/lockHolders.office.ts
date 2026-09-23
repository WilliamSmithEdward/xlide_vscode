// Who holds a file, asked of Restart Manager for real (src/fileLockHolders.ts).

import * as cp from 'child_process';
import { afterAll, describe, expect, it } from 'vitest';
import { describeLockHolders, findFileLockHolders } from '../../src/fileLockHolders';
import { buildHostOpenScript } from '../../src/officeHostLauncher';
import { hostIsFree, ps, releaseHost, scratchCopy, sleep } from './officeHarness';

/** A PowerShell process holding the file with no sharing, until it is told to let go. */
function holdFile(filePath: string): { pid: Promise<number>; release: () => Promise<void> } {
    const script = [
        `$fs = [System.IO.File]::Open('${filePath.replace(/'/g, "''")}', 'Open', 'ReadWrite', 'None')`,
        '[Console]::Out.WriteLine("holding " + $PID)',
        '[void][Console]::In.ReadLine()',
        '$fs.Close()',
    ].join('; ');
    const child = cp.spawn('powershell.exe', ['-NoProfile', '-Command', script], { windowsHide: true });
    const pid = new Promise<number>((resolve, reject) => {
        child.stdout.on('data', (chunk: Buffer) => {
            const match = /holding (\d+)/.exec(chunk.toString());
            if (match) {
                resolve(Number(match[1]));
            }
        });
        child.on('error', reject);
    });
    const release = (): Promise<void> => new Promise((resolve) => {
        child.on('close', () => resolve());
        child.stdin.end('\n');
    });
    return { pid, release };
}

describe('Restart Manager names who holds a file', () => {
    it('names a process holding it, and nobody once it lets go', async () => {
        const file = scratchCopy('NoVbaFixture.xlsm', 'holders-plain');
        expect(await findFileLockHolders(file)).toEqual([]);

        const holder = holdFile(file);
        const pid = await holder.pid;
        const holders = await findFileLockHolders(file);
        expect(holders).toEqual([{ pid, appName: 'Windows PowerShell', image: 'powershell.exe' }]);
        expect(describeLockHolders(holders!)).toBe(`Windows PowerShell (powershell.exe, process ${pid})`);

        await holder.release();
        expect(await findFileLockHolders(file)).toEqual([]);
    });
});

const wordFree = await hostIsFree('word');

describe.runIf(wordFree)('Restart Manager and Word', () => {
    afterAll(async () => {
        await releaseHost('word');
    });

    it('names Word holding a document it has open only read-only', async () => {
        // The question that started this: a file open read-only in Word, and
        // still XLIDE could not write to it. Word locks it anyway.
        const file = scratchCopy('WordFixture.docm', 'holders-word');
        const open = await ps(buildHostOpenScript({ host: 'word', filePath: file, attachToRunning: true, readOnly: true }));
        expect(open.out).toContain('XLIDE_OPEN|opened');
        const wordPid = Number((await ps('(Get-Process -Name WINWORD | Select-Object -First 1).Id')).out[0]);
        await sleep(500);
        const holders = await findFileLockHolders(file);
        expect(holders?.map((holder) => [holder.pid, holder.image])).toEqual([[wordPid, 'WINWORD.EXE']]);
        expect(describeLockHolders(holders!)).toMatch(/^Microsoft Word \(WINWORD\.EXE, process \d+\)$/);
    });
});
