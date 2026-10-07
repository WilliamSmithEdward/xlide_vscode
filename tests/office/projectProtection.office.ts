// Real Office checks for the native locked corpus and protected host copies.
// The automation baseline records VBE add-in state explicitly: an installed
// add-in is not necessarily loaded in an embedding-mode COM host.
import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import { openMacroContainer } from '../../src/vba/macroContainer';
import { VbaProject } from '../../src/vba/vbaProject';
import { hostIsFree, ps, q, scratchCopy } from './officeHarness';
import { OFFICE_HOST_APPS, type OfficeHostApp } from '../../src/officeHostApps';

const CASES: Array<[OfficeHostApp, string]> = [
    ['excel', 'PasswordProtectedFixture.xlsm'],
    ['excel', 'PasswordProtectedFormFixture.xlsm'],
    ['word', 'PasswordProtectedWordFixture.docm'],
    ['powerpoint', 'PasswordProtectedPowerPointFixture.pptm'],
    ['access', 'PasswordProtectedAccessFixture.accdb'],
];
describe('Protected VBA password fixtures in real Office', () => {
    it.each(CASES)('%s recognizes %s as locked for viewing', async (host, fixture, context) => {
        if (!await hostIsFree(host)) { context.skip(); }
        const file = scratchCopy(fixture, `password-${host}-${fixture}`);
        const before = fs.readFileSync(file);
        const info = OFFICE_HOST_APPS[host];
        const open = host === 'excel' ? `$document = $app.Workbooks.Open(${q(file)}, 0, $true)`
            : host === 'word' ? `$document = $app.Documents.Open(${q(file)}, $false, $true)`
            : host === 'powerpoint' ? `$document = $app.Presentations.Open(${q(file)}, -1, 0, 0)`
            : `$app.OpenCurrentDatabase(${q(file)}, $false)`;
        const project = host === 'access' ? '$project = $app.VBE.ActiveVBProject' : '$project = $document.VBProject';
        const close = host === 'excel' ? '$document.Close($false)' : host === 'word' ? '$document.Close(0)' : '$document.Close()';
        const quit = host === 'access' ? '$app.Quit(2)' : '$app.Quit()';
        const alerts = host === 'excel' ? '$app.DisplayAlerts = $false' : host === 'word' ? '$app.DisplayAlerts = 0' : host === 'powerpoint' ? '$app.DisplayAlerts = 1' : '';
        const result = await ps(`
$ErrorActionPreference = 'Stop'
$app = $null; $document = $null; $project = $null; $owned = $false; $vbe = $null; $addins = $null
try {
    if (@(Get-Process -Name ${info.processName} -ErrorAction SilentlyContinue).Count -ne 0) { "owned-instance=no"; return }
    $app = New-Object -ComObject ${info.progId}
    # The host was absent and COM has created one process. Refuse a race.
    if (@(Get-Process -Name ${info.processName} -ErrorAction SilentlyContinue).Count -ne 1) { "owned-instance=no"; return }
    $owned = $true
    ${alerts}
    $app.AutomationSecurity = 3
    ${open}
    ${project}
    "protection=$([int]$project.Protection)"
    $vbe = $app.VBE; $addins = $vbe.Addins
    $found = $false
    foreach ($addin in $addins) {
        try { if ($addin.ProgId -eq 'Xlide.VbeAddIn') { $found = $true; "xlide-vbide-connected=$([bool]$addin.Connect)" } }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($addin) }
    }
    if (-not $found) { "xlide-vbide-connected=not-loaded" }
} finally {
    if ($addins) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($addins) }
    if ($vbe) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($vbe) }
    if ($project) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($project) }
    if ($document) { ${close}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($app) { if ($owned) { ${quit} }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
    $addins = $null; $vbe = $null; $project = $null; $document = $null; $app = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
`);
        expect(result.code, JSON.stringify(result)).toBe(0);
        if (result.out.includes('owned-instance=no')) { context.skip(); }
        expect(result.out).toContain('protection=1');
        expect(result.out.some(line => line.startsWith('xlide-vbide-connected='))).toBe(true);
        console.log(`${host}: ${result.out.join('; ')}`);
        if (host === 'access') {
            // Access updates database metadata when opened, even without saving code.
            const project = VbaProject.parse(openMacroContainer(fs.readFileSync(file)).vbaCfb());
            expect(project.protection.verify('Test66')).toBe(true);
        } else {
            expect(fs.readFileSync(file)).toEqual(before);
        }
    });
});
