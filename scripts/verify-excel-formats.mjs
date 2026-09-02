#!/usr/bin/env node
// Verify every creatable file format against real Excel.
//
// XLIDE builds .xlsm / .xlsb / .xlam from baked-in templates without Office
// involved, so the thing worth proving is that Excel agrees with what we
// claim to have produced. The offline tests pin the container markers; this
// harness closes the loop by asking Excel itself.
//
// For each format it creates a file through the real service, writes a
// sentinel macro into it, then opens it in Excel and RUNS that macro. The
// macro reports ThisWorkbook.IsAddin from inside the file, so an .xlam that
// is merely a renamed workbook cannot pass.
//
// Like the VBE oracle, this is a developer check and not part of CI: it needs
// Excel installed, and it exits 0 with a skip when Excel is absent. It has no
// dependencies beyond Node built-ins, PowerShell, and esbuild (already a dev
// dependency) to bundle the TypeScript service.

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const WORKER = path.join(ROOT, 'scripts', 'verify_excel_formats_worker.ps1');
const TEMPLATE_DIR = path.join(ROOT, 'assets', 'templates');

// Module and procedure are deliberately named differently: Application.Run
// resolves a bare name to the module when the two match.
const PROBE_MODULE = 'XlideFormatProbe';
const PROBE_PROC = 'ReportFormat';
const PROBE_SOURCE = [
    'Option Explicit',
    '',
    `Public Function ${PROBE_PROC}() As String`,
    `    ${PROBE_PROC} = "XLIDE-OK|" & ThisWorkbook.Name & "|IsAddin=" & CStr(ThisWorkbook.IsAddin)`,
    'End Function',
    '',
].join('\r\n');

// extension -> [template, expected IsAddin, expected Excel FileFormat]
// 52 xlOpenXMLWorkbookMacroEnabled, 50 xlExcel12 (binary), 55 xlOpenXMLAddIn.
const FORMATS = [
    ['.xlsm', 'blank.xlsm', false, 52],
    ['.xlsb', 'blank.xlsb', false, 50],
    ['.xlam', 'blank.xlam', true, 55],
];

function excelComRegistered() {
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
        '$type = [type]::GetTypeFromProgID("Excel.Application"); if ($null -eq $type) { exit 2 }',
    ], { encoding: 'utf8', timeout: 15_000 });
    return !completed.error && completed.status === 0;
}

/** Bundle the standalone workbook service so plain node can call it. */
async function loadProjectService(workDir) {
    const esbuild = await import('esbuild');
    const bundle = path.join(workDir, 'projectService.mjs');
    await esbuild.build({
        entryPoints: [path.join(ROOT, 'src', 'vba', 'projectService.ts')],
        outfile: bundle,
        bundle: true,
        format: 'esm',
        platform: 'node',
        logLevel: 'silent',
    });
    return import(new URL(`file://${bundle.replace(/\\/g, '/')}`));
}

function main(svc, workDir) {
    const files = [];
    for (const [extension, template] of FORMATS) {
        const target = path.join(workDir, `Probe${extension}`);
        svc.createProject(target, path.join(TEMPLATE_DIR, template));
        svc.writeModule(target, PROBE_MODULE, PROBE_SOURCE, 'standard');
        const back = svc.readModule(target, PROBE_MODULE, false).source;
        if (!back.includes(`${PROBE_PROC} = "XLIDE-OK|"`)) {
            throw new Error(`${extension}: module did not read back from the file we wrote`);
        }
        files.push(target);
    }

    const manifest = path.join(workDir, 'files.txt');
    fs.writeFileSync(manifest, files.join('\n'), 'utf8');
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', WORKER,
        '-Manifest', manifest, '-Module', PROBE_MODULE, '-Procedure', PROBE_PROC,
    ], { encoding: 'utf8', timeout: 300_000 });
    if (completed.error) { throw completed.error; }
    if (completed.stderr?.trim()) { process.stderr.write(completed.stderr); }

    const results = JSON.parse(fs.readFileSync(path.join(workDir, 'results.json'), 'utf8')
        .replace(/^﻿/, ''));
    let failed = 0;
    for (const [extension, , expectAddin, expectFormat] of FORMATS) {
        const name = `Probe${extension}`;
        const row = results.find((entry) => entry.name === name);
        const problems = [];
        if (!row) {
            problems.push('Excel produced no result');
        } else {
            if (row.error) { problems.push(row.error); }
            if (row.isAddin !== expectAddin) {
                problems.push(`IsAddin ${row.isAddin} (expected ${expectAddin})`);
            }
            if (row.fileFormat !== expectFormat) {
                problems.push(`FileFormat ${row.fileFormat} (expected ${expectFormat})`);
            }
            const wanted = `XLIDE-OK|${name}|IsAddin=${expectAddin ? 'True' : 'False'}`;
            if (row.macro !== wanted) {
                problems.push(`macro returned ${JSON.stringify(row.macro)} (expected ${JSON.stringify(wanted)})`);
            }
        }
        if (problems.length) {
            failed++;
            console.log(`FAIL ${name}: ${problems.join('; ')}`);
        } else {
            console.log(`ok   ${name}  IsAddin=${row.isAddin}  FileFormat=${row.fileFormat}  macro => ${row.macro}`);
        }
    }
    return failed;
}

if (!excelComRegistered()) {
    console.log('skip: Excel COM is not registered on this machine.');
    process.exit(0);
}

const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-formats-'));
let failed = 1;
try {
    failed = main(await loadProjectService(workDir), workDir);
} finally {
    fs.rmSync(workDir, { recursive: true, force: true });
}
process.exit(failed ? 1 : 0);
