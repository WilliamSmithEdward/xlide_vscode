// F5 on a procedure in an Access database Access does not trust (issue
// #1531). XLIDE never changes the Trust Center, a Trusted Location or a
// file's zone, and never answers Access's prompts: it reads
// CurrentProject.IsTrusted after the open and says why nothing ran.
//
// Runs only when no Access is running, so the instance XLIDE attaches to is
// one this check started. Each case's procedure writes a marker file, which
// is how the check knows whether it ran. Only an Access whose database lies
// under the scratch folder is quit, and no process is killed.

import * as fs from 'fs';
import * as path from 'path';
import { afterEach, describe, expect, it } from 'vitest';
import { buildAccessMacroLaunchScript } from '../../src/officeHostLauncher';
import { writeModule } from '../../src/vba/projectService';
import { hostIsFree, ps, psFile, q, SCRATCH, scratchCopy, waitForExit } from './officeHarness';

const accessFree = await hostIsFree('access');

/** A scratch database whose XlideTrustProbe writes `marker` when it runs. */
function probeDatabase(label: string): { file: string; marker: string } {
    const file = scratchCopy('AccessFixture.accdb', label);
    const marker = path.join(path.dirname(file), 'ran.txt');
    writeModule(file, 'XlideTrust', [
        'Option Explicit',
        'Public Sub XlideTrustProbe()',
        `    Open "${marker}" For Output As #1`,
        '    Print #1, "ran"',
        '    Close #1',
        'End Sub',
        '',
    ].join('\r\n'), 'standard');
    return { file, marker };
}

/** Quits the Access this check left running, if its database is the check's own. */
async function releaseAccess(): Promise<void> {
    await ps([
        `$root = ${q(SCRATCH)}`,
        '$app = $null',
        'try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application") } catch { }',
        'if ($app) {',
        '  $open = ""',
        '  try { $open = $app.CurrentProject.FullName } catch { }',
        '  if ($open -like ($root + "*")) { try { $app.CloseCurrentDatabase() } catch { }; $app.Quit() }',
        '}',
        '$app = $null',
        '[GC]::Collect()',
    ].join('\n'));
    await waitForExit('access', 60_000);
}

async function pressF5(file: string, attachToRunning: boolean): Promise<string[]> {
    const result = await psFile(buildAccessMacroLaunchScript(file, 'XlideTrustProbe', attachToRunning), 'access-trust-f5.ps1');
    return [...result.out, ...result.err];
}

describe.runIf(accessFree)('Access F5 and trust (issue #1531)', () => {
    afterEach(async () => { await releaseAccess(); });

    it('says why nothing ran when Access opened the database with its code disabled', async (context) => {
        const { file, marker } = probeDatabase('access-trust-disabled');
        // A machine whose Trust Center enables all macros trusts every
        // database not marked as downloaded, and the check never changes that.
        const measured = await psFile([
            '$app = New-Object -ComObject Access.Application',
            `try { $app.OpenCurrentDatabase(${q(file)}); "trusted=" + $app.CurrentProject.IsTrusted; $app.CloseCurrentDatabase() } finally { $app.Quit(); $app = $null; [GC]::Collect() }`,
        ].join('\r\n'), 'access-trust-measure.ps1');
        await waitForExit('access', 60_000);
        if (measured.out.includes('trusted=True')) {
            context.skip('Access on this machine trusts a database outside a trusted location (Enable all macros)');
        }
        const output = (await pressF5(file, false)).join('\n');
        expect(fs.existsSync(marker)).toBe(false);
        expect(output).toContain('XLIDE_MACRO_ERROR|RUN_FAILED|');
        expect(output).toContain('Enable Content');
        expect(output).toContain('Trusted Locations');
        expect(output).not.toContain('0x800A9D9F');
    });

    it('tells a downloaded database apart, and says to unblock it', async () => {
        const { file, marker } = probeDatabase('access-trust-downloaded');
        // Zone 3 is the Internet zone: what a browser marks a download with.
        fs.writeFileSync(`${file}:Zone.Identifier`, '[ZoneTransfer]\r\nZoneId=3\r\n');
        const output = (await pressF5(file, false)).join('\n');
        expect(fs.existsSync(marker)).toBe(false);
        expect(output).toContain('XLIDE_MACRO_ERROR|RUN_FAILED|');
        expect(output).toContain('Unblock');
    });

    it('runs the procedure in a database Access trusts', async () => {
        const { file, marker } = probeDatabase('access-trust-trusted');
        // msoAutomationSecurityLow (1), from officeReferenceConstants.ts: a
        // database opened this way is trusted, and F5 attaches to it.
        const opened = await psFile([
            '$ErrorActionPreference = "Stop"',
            '$app = New-Object -ComObject Access.Application',
            '$app.AutomationSecurity = 1',
            `$app.OpenCurrentDatabase(${q(file)})`,
            '$app.UserControl = $true',
            '"trusted=" + $app.CurrentProject.IsTrusted',
        ].join('\r\n'), 'access-trust-open.ps1');
        expect(opened.out).toContain('trusted=True');
        const output = (await pressF5(file, true)).join('\n');
        expect(output).toContain('XLIDE_MACRO_OK');
        expect(fs.readFileSync(marker, 'utf8').trim()).toBe('ran');
    });
});
