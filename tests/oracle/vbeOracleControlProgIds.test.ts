// A control's ProgID in an oracle case names a class registered on the
// machine, and Office creates it when the worker adds the control: in a
// compile case too, where none of the case's VBA runs. So the worker takes
// only the ProgIDs listed here, and refuses any other before it asks the host
// for anything. This file runs without Office, so a pull request whose case
// names a new ProgID fails CI before anyone runs the oracle.
//
// To add a ProgID: measure the control in every host that has it, add it to
// the worker's list ($DesignerControlInterfaces with its read-back type, or
// $DocumentControlProgIds) and to the matching list below, in the same pull
// request as the case that needs it.

import { spawnSync } from 'node:child_process';
import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';
import { oracleModuleInput } from '../helpers/oracleModuleInput';

/** What a UserForm's `designerControls` may name: the keys of the worker's `$DesignerControlInterfaces`. */
const DESIGNER_CONTROL_PROG_IDS = ['Forms.CommandButton.1', 'Forms.ComboBox.1'];
/** What a case's `documentControls` may name: the worker's `$DocumentControlProgIds`. */
const DOCUMENT_CONTROL_PROG_IDS = ['Forms.CommandButton.1', 'Forms.CheckBox.1', 'Forms.TextBox.1'];

const ORACLE = path.join(__dirname, '..', '..', 'syntax_corpus', 'oracle');
const worker = fs.readFileSync(path.join(ORACLE, 'vbe_oracle_worker.ps1'), 'utf8').replaceAll('\r\n', '\n');

/** A function of the worker, from its `function` line to the line before the next one. */
function body(name: string): string {
    const start = worker.search(new RegExp(`^function ${name}\\b`, 'm'));
    expect(start, name).toBeGreaterThanOrEqual(0);
    return worker.slice(start, worker.indexOf('\nfunction ', start + 1));
}

const designerTable = /^\$DesignerControlInterfaces = @\{\n([\s\S]*?)\n\}$/m.exec(worker)?.[0] ?? '';
const documentList = /^\$DocumentControlProgIds = @\(\n([\s\S]*?)\n\)$/m.exec(worker)?.[0] ?? '';

const sorted = (ids: readonly string[]): string[] => [...ids].sort();
const allowed = (list: readonly string[], progId: unknown): boolean =>
    typeof progId === 'string' && list.some((known) => known.toLowerCase() === progId.toLowerCase());

describe('the control ProgIDs the VBE oracle may create', () => {
    it('lists the same ProgIDs here and in the worker', () => {
        const designer = designerTable.split('\n').slice(1, -1).map((line) => /^ {4}"([^"]+)" = "\w+"$/.exec(line)?.[1]);
        const documents = documentList.split('\n').slice(1, -1).map((line) => /^ {4}"([^"]+)"$/.exec(line)?.[1]);
        // Every line of each list is an entry this test read.
        expect(designer).not.toContain(undefined);
        expect(documents).not.toContain(undefined);
        expect(sorted(designer as string[])).toEqual(sorted(DESIGNER_CONTROL_PROG_IDS));
        expect(sorted(documents as string[])).toEqual(sorted(DOCUMENT_CONTROL_PROG_IDS));
        // Each list is assigned once, so nothing later in the worker widens it.
        expect(worker.match(/\$DesignerControlInterfaces\s*(\[[^\]]*\]\s*)?[+]?=|\$DesignerControlInterfaces\.(Add|Set_Item)\b/g)).toHaveLength(1);
        expect(worker.match(/\$DocumentControlProgIds\s*(\[[^\]]*\]\s*)?[+]?=/g)).toHaveLength(1);
        // The analyzer's adapter knows a type for every one of them.
        for (const progId of [...DESIGNER_CONTROL_PROG_IDS, ...DOCUMENT_CONTROL_PROG_IDS]) {
            expect(() => oracleModuleInput({ name: 'F', type: 'userform', source: '', designerControls: [{ name: 'c', progId }] })).not.toThrow();
        }
    });

    it('finds only listed ProgIDs in the cases, and only where the worker checks them', () => {
        const corpus: { cases: Array<{ id: string }> } = JSON.parse(fs.readFileSync(path.join(ORACLE, 'vbe_oracle_cases.json'), 'utf8'));
        const problems: string[] = [];
        const lists: Array<[RegExp, readonly string[]]> = [
            [/^documentControls\.\d+\.progId$/, DOCUMENT_CONTROL_PROG_IDS],
            [/^(projectReference\.)?modules\.\d+\.designerControls\.\d+\.progId$/, DESIGNER_CONTROL_PROG_IDS],
        ];
        const fields = [/^documentControls$/, /^(projectReference\.)?modules\.\d+\.designerControls$/];
        let seen = 0;
        const walk = (node: unknown, at: string, id: string): void => {
            if (node === null || typeof node !== 'object') {
                return;
            }
            for (const [key, value] of Object.entries(node)) {
                const here = at === '' ? key : `${at}.${key}`;
                // The worker reads properties without regard to case.
                const name = key.toLowerCase();
                if (name === 'progid') {
                    seen += 1;
                    const list = lists.find(([where]) => where.test(here))?.[1];
                    if (list === undefined) {
                        problems.push(`${id}: a ProgID at ${here}, which the worker does not check`);
                    } else if (!allowed(list, value)) {
                        problems.push(`${id}: ${here} is ${JSON.stringify(value)}, which is not listed`);
                    }
                }
                if ((name === 'documentcontrols' || name === 'designercontrols') && !fields.some((where) => where.test(here))) {
                    problems.push(`${id}: ${here} is not a place the worker checks`);
                }
                if ((name === 'documentcontrols' || name === 'designercontrols') && Array.isArray(value)) {
                    value.forEach((control, i) => {
                        if (control === null || typeof control !== 'object' || typeof (control as { progId?: unknown }).progId !== 'string') {
                            problems.push(`${id}: ${here}.${i} has no ProgID`);
                        }
                    });
                }
                walk(value, here, id);
            }
        };
        for (const kase of corpus.cases) {
            walk(kase, '', kase.id);
        }
        expect(problems).toEqual([]);
        expect(seen).toBeGreaterThan(0);
    });

    it('checks every ProgID of a case before the host starts', () => {
        // Without its comment lines, one of which names New-ProjectReference.
        const run = body('Invoke-OracleCase').replace(/^\s*#.*\n/gm, '');
        const check = run.indexOf('Assert-ApprovedControls $case');
        expect(check).toBeGreaterThan(0);
        for (const first of ['$script:app', 'Start-OracleHost', 'Stop-OracleHost', 'Add-DocumentControls', 'Add-OracleComponent', 'New-ProjectReference']) {
            expect(run.indexOf(first), first).toBeGreaterThan(check);
        }
        // Inside the try, so a refusal is a worker_error with its message.
        expect(run.slice(run.indexOf('    try {'), check)).toMatch(/^ {4}try \{\n\s*Set-Stage "check_controls"\n\s*$/);

        const all = body('Assert-ApprovedControls');
        expect(all).toContain('Assert-ApprovedProgId ([string]$control.progId) $DocumentControlProgIds "documentControls"');
        expect(all).toContain('Assert-ApprovedProgId ([string]$control.progId) $DesignerControlInterfaces.Keys "designerControls"');
        expect(all).toContain('@($Case.projectReference.modules)');
        const one = body('Assert-ApprovedProgId');
        expect(one).toContain('[string]::Equals($ProgId, [string]$known, [StringComparison]::OrdinalIgnoreCase)');
        expect(one.match(/throw "/g)).toHaveLength(1);
    });

    it('checks them again where each control is made, before the first host call', () => {
        const designer = body('Add-DesignerControls');
        const designerCheck = designer.indexOf('Assert-ApprovedProgId ([string]$control.progId) $DesignerControlInterfaces.Keys "designerControls"');
        expect(designerCheck).toBeGreaterThan(0);
        expect(designer.indexOf('$Component.')).toBeGreaterThan(designerCheck);
        expect(designer.indexOf('$designer')).toBeGreaterThan(designerCheck);
        // The one call that makes a control is given the checked value.
        expect(designer.match(/\.Add\(/g)).toHaveLength(1);
        expect(designer).toContain('$designer.Controls.Add([string]$control.progId, [string]$control.name, $true)');

        const documents = body('Add-DocumentControls');
        const documentCheck = documents.indexOf('Assert-ApprovedProgId ([string]$control.progId) $DocumentControlProgIds "documentControls"');
        expect(documentCheck).toBeGreaterThan(0);
        expect(documents.indexOf('$Container.')).toBeGreaterThan(documentCheck);
        expect(documents.indexOf('switch ($script:hostName)')).toBeGreaterThan(documentCheck);
        expect(documents).toContain('$progId = [string]$control.progId');
        // Every call that makes a control is given the checked value, in each
        // host that has them; Access has no arm and is refused as before.
        // ($held is the worker's own list of references to release, and
        // $slides.Add(1, 12) the blank slide a control goes on.)
        expect([...documents.matchAll(/(?<!\$held)\.(AddOLEControl|AddOLEObject|Add)\(([^)]*)\)/g)].map((m) => `${m[1]}(${m[2]})`).sort()).toEqual([
            'Add($progId)', 'Add(1, 12)',
            'AddOLEControl($progId)', 'AddOLEControl($progId)', 'AddOLEControl($progId)',
            'AddOLEObject(10, 10, 120, 30, $progId)',
        ]);
        for (const host of Object.keys(OFFICE_HOST_APPS)) {
            expect(documents.includes(`"${host}" {`), host).toBe(host !== 'access');
        }
        expect(documents).toContain('default { throw "documentControls are measured in Word, Excel and PowerPoint." }');
    });

    it.skipIf(process.platform !== 'win32')('refuses an unlisted ProgID without touching the host, in every host', () => {
        const control = (progId: unknown, name = 'CommandButton1'): object => ({ name, progId });
        const form = (...progIds: unknown[]): object => ({ name: 'Form1', type: 'userform', source: '', designerControls: progIds.map((p, i) => control(p, `c${i}`)) });
        const cases: Array<[string, object, 'ok' | 'documentControls' | 'designerControls']> = [
            ['no controls', { id: 'a', source: '' }, 'ok'],
            ['every listed ProgID', {
                documentControls: DOCUMENT_CONTROL_PROG_IDS.map((p) => control(p)),
                modules: [form(...DESIGNER_CONTROL_PROG_IDS)],
                projectReference: { name: 'Lib', modules: [form(...DESIGNER_CONTROL_PROG_IDS)] },
            }, 'ok'],
            ['another case of the letters', { documentControls: [control('forms.COMMANDBUTTON.1')], modules: [form('FORMS.combobox.1')] }, 'ok'],
            ['an empty designerControls', { modules: [form()] }, 'ok'],
            ['an unlisted class on the document', { documentControls: [control('Scripting.FileSystemObject')] }, 'documentControls'],
            ['an unlisted class after a listed one', { documentControls: [control('Forms.TextBox.1'), control('WScript.Shell')] }, 'documentControls'],
            ['a designer ProgID on the document', { documentControls: [control('Forms.ComboBox.1')] }, 'documentControls'],
            ['a document ProgID in a designer', { modules: [form('Forms.CheckBox.1')] }, 'designerControls'],
            ['an unlisted class in a designer, after a listed one', { modules: [form('Forms.CommandButton.1', 'Shell.Application')] }, 'designerControls'],
            ['an unlisted class in a later module', { modules: [{ name: 'M', type: 'standard', source: '' }, form('MSComctlLib.TreeCtrl.2')] }, 'designerControls'],
            ['an unlisted class in the referenced project', { projectReference: { name: 'Lib', modules: [form('Forms.Image.1')] } }, 'designerControls'],
            ['no ProgID at all', { documentControls: [{ name: 'CommandButton1' }] }, 'documentControls'],
            ['a trailing space', { documentControls: [control('Forms.TextBox.1 ')] }, 'documentControls'],
            ['a zero-width character', { documentControls: [control(`Forms.Command${String.fromCharCode(0x200d)}Button.1`)] }, 'documentControls'],
            ['a Kelvin sign for the k', { documentControls: [control(`Forms.Chec${String.fromCharCode(0x212a)}Box.1`)] }, 'documentControls'],
            ['a class ID', { documentControls: [control('{D7053240-CE69-11CD-A777-00DD01143C57}')] }, 'documentControls'],
        ];
        const functions = [designerTable, documentList, ...['Assert-ApprovedProgId', 'Assert-ApprovedControls', 'Add-DesignerControls', 'Add-DocumentControls'].map(body)].join('\n');
        // The host objects are stand-ins that count every use: no Office
        // application is started.
        const script = String.raw`
$ErrorActionPreference = 'Stop'
. ([scriptblock]::Create($env:XLIDE_CONTROL_CHECKS))
function Remove-ComObjectReference($Reference) { }
$script:touched = 0
function New-StandIn([string[]]$Members) {
    $standIn = New-Object psobject
    foreach ($member in $Members) {
        $standIn | Add-Member -MemberType ScriptProperty -Name $member -Value { $script:touched += 1; throw 'The host was touched.' }
    }
    return $standIn
}
function Invoke-Check([scriptblock]$Check) {
    $script:touched = 0
    $said = try { & $Check; 'ok' } catch { 'refused: ' + $_.Exception.Message }
    return [pscustomobject]@{ said = $said; touched = $script:touched }
}
$cases = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($env:XLIDE_CONTROL_CASES)) | ConvertFrom-Json
$whole = foreach ($case in $cases) { Invoke-Check { Assert-ApprovedControls $case } }
$bad = [pscustomobject]@{ name = 'CommandButton2'; progId = 'Scripting.FileSystemObject' }
$good = [pscustomobject]@{ name = 'CommandButton1'; progId = 'Forms.CommandButton.1' }
$made = foreach ($hostName in @('excel', 'word', 'powerpoint', 'access')) {
    $script:hostName = $hostName
    $container = New-StandIn @('InlineShapes', 'Shapes', 'Sections', 'Worksheets', 'Slides')
    $component = New-StandIn @('Designer')
    $document = Invoke-Check { Add-DocumentControls $container ([pscustomobject]@{ documentControls = @($good, $bad) }) }
    $designer = Invoke-Check { Add-DesignerControls $component ([pscustomobject]@{ name = 'Form1'; designerControls = @($good, $bad) }) }
    $listed = Invoke-Check { Add-DocumentControls $container ([pscustomobject]@{ documentControls = @($good) }) }
    [pscustomobject]@{ host = $hostName; document = $document; designer = $designer; listed = $listed }
}
[pscustomobject]@{ whole = @($whole); made = @($made) } | ConvertTo-Json -Depth 5 -Compress
`;
        const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
            encoding: 'utf8',
            timeout: 60_000,
            env: {
                ...process.env,
                XLIDE_CONTROL_CHECKS: functions,
                XLIDE_CONTROL_CASES: Buffer.from(JSON.stringify(cases.map(([, kase]) => kase)), 'utf8').toString('base64'),
            },
        });
        expect(run.error).toBeUndefined();
        expect(run.status, run.stderr).toBe(0);
        const seen: {
            whole: Array<{ said: string; touched: number }>;
            made: Array<Record<'document' | 'designer' | 'listed', { said: string; touched: number }> & { host: string }>;
        } = JSON.parse(run.stdout.trim());

        expect(seen.whole).toHaveLength(cases.length);
        cases.forEach(([label, , verdict], i) => {
            if (verdict === 'ok') {
                expect(seen.whole[i].said, label).toBe('ok');
            } else {
                expect(seen.whole[i].said, label).toMatch(new RegExp(`^refused: The ${verdict} ProgID '.*' is not approved, so no control is made\\. Approved: `));
            }
        });

        expect(seen.made.map((made) => made.host)).toEqual(Object.keys(OFFICE_HOST_APPS));
        for (const made of seen.made) {
            // The unlisted control is second, and the listed one before it is not made either.
            expect(made.document, made.host).toEqual({
                said: "refused: The documentControls ProgID 'Scripting.FileSystemObject' is not approved, so no control is made. Approved: Forms.CheckBox.1, Forms.CommandButton.1, Forms.TextBox.1.",
                touched: 0,
            });
            expect(made.designer, made.host).toEqual({
                said: "refused: The designerControls ProgID 'Scripting.FileSystemObject' is not approved, so no control is made. Approved: Forms.ComboBox.1, Forms.CommandButton.1.",
                touched: 0,
            });
            // A listed ProgID goes on to the host, except in Access, which has no such controls.
            if (made.host === 'access') {
                expect(made.listed).toEqual({ said: 'refused: documentControls are measured in Word, Excel and PowerPoint.', touched: 0 });
            } else {
                expect(made.listed.touched, made.host).toBe(1);
            }
        }
    }, 70_000);
});
