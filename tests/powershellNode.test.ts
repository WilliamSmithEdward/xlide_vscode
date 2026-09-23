import { EventEmitter } from 'events';
import { describe, expect, it, vi } from 'vitest';

const spawned = vi.hoisted(() => ({ calls: [] as Array<{ file: string; args: string[] }> }));

vi.mock('child_process', () => ({
    spawn: vi.fn((file: string, args: string[]) => {
        spawned.calls.push({ file, args });
        const child = new EventEmitter() as EventEmitter & { stdout: EventEmitter; stderr: EventEmitter; pid: number; kill: () => void };
        child.stdout = new EventEmitter();
        child.stderr = new EventEmitter();
        child.pid = 1;
        child.kill = () => undefined;
        setImmediate(() => child.emit('close', 0, null));
        return child;
    }),
}));

import { encodedCommandArgs, runPowerShell } from '../src/util/powershellNode';

const decode = (base64: string): string => Buffer.from(base64, 'base64').toString('utf16le');

describe('how a script reaches PowerShell', () => {
    it('travels whole as -EncodedCommand, line breaks and all', () => {
        const script = 'if ($x) {\n  "a"\n}\nelse {\n  "b"\n}';
        const [flag, payload] = encodedCommandArgs(script);
        expect(flag).toBe('-EncodedCommand');
        expect(decode(payload)).toBe(`[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n${script}`);
    });

    it('carries text outside ASCII intact', () => {
        const script = "Write-Output 'Bob''s \u00e9 \u0416 report'";
        expect(decode(encodedCommandArgs(script)[1]).endsWith(script)).toBe(true);
    });

    it('is what runPowerShell sends for a script, and a script file still goes by -File', async () => {
        spawned.calls.length = 0;
        await runPowerShell({ script: '"one"\n"two"' }).result;
        await runPowerShell({ args: ['-File', 'C:\\host\\run.ps1'] }).result;
        const [scriptRun, fileRun] = spawned.calls;
        expect(scriptRun.file).toBe('powershell.exe');
        expect(scriptRun.args.slice(0, 4)).toEqual(['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand']);
        expect(decode(scriptRun.args[4])).toContain('"one"\n"two"');
        expect(scriptRun.args).not.toContain('-Command');
        expect(fileRun.args).toEqual(['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'C:\\host\\run.ps1']);
    });
});
