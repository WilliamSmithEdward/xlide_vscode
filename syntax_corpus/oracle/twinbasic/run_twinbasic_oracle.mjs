#!/usr/bin/env node
// The twinBASIC oracle: does twinBASIC's compiler accept a VBA/VB6 snippet?
//
// A developer audit tool, Windows only, never part of `npm test`. twinBASIC is
// a superset of VB6 with published incompatibilities, so its verdict is typed
// evidence about twinBASIC and inferred evidence about VB6; the harness
// records it as such and never as "VB6-accepted". Shaped like
// ../run_excel_vbe_oracle.mjs: one case at a time, a bounded run, a kill that
// reaches only what the run spawned, and only `accepted`/`rejected` as
// evidence.
//
// How a case is asked (measured on twinBASIC BETA 983, see README.md):
//   1. stage a project folder: Settings (from Settings.template.json), one
//      `.twin` per module wrapping the VB6 source in `Module`/`Class`, and a
//      Startup module whose Main reaches every procedure the case declares,
//      because twinBASIC's build only fails on code it links;
//   2. `bin\twinBASIC_win64.exe import <case.twinproj> <folder>` (headless);
//   3. `twinBASIC.exe <case.twinproj> --buildAndExit64`, watched:
//        exits and Build\ holds the EXE          -> accepted
//        stays up past the stall window, Build\ exists, no EXE
//                                                -> rejected (the compiler
//           reported errors; the IDE stalls on them by design and prints
//           nothing outside its own panels)
//        anything else                           -> an infrastructure outcome
//   4. the IDE process tree is killed by PID.
// Every batch runs an accept control and a reject control before and after
// the cases; rejections count only when all four controls passed.
//
// Usage: node syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs [options]
//   --cases <path>        fixture document (default: ../vbe_oracle_cases.json,
//                         the VBE corpus, for the parity report)
//   --case <id>           run only this case; repeatable
//   --limit <n>           run at most n cases
//   --offset <n>          skip the first n cases
//   --append              with --report: keep the report's earlier evidence
//                         and skip the cases it already holds, so a long
//                         batch can be run in chunks
//   --twinbasic <dir>     the IDE folder (twinBASIC.exe, bin\twinBASIC_win64.exe);
//                         default: $XLIDE_TWINBASIC_DIR
//   --timeout <seconds>   per-build ceiling (default 60)
//   --stall <seconds>     how long Build\ may stay empty with the IDE up
//                         before the build counts as rejected (default 4)
//   --parallel <n>        IDE instances at once (default 1); each case has
//                         its own folder and process tree
//   --stagger <ms>        delay between instance starts (default 1500)
//   --excel               also reference the Excel and Office type libraries
//                         registered on this machine, as the VBE corpus had
//                         them; off by default so language cases need no Excel
//   --startup <seconds>   how long an instance may take to start the build
//                         before it is killed and the case retried once
//                         (default 20)
//   --no-controls         skip the controls (debugging only; results are
//                         then observations, not evidence)
//   --keep                keep the staged folders
//   --json                machine-readable output
//   --report <path>       write results JSON (parity_results.json)
//   --parity-doc <path>   write the parity matrix as Markdown
//   --strict              exit non-zero on expectation mismatches

import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_CASES = path.join(HERE, '..', 'vbe_oracle_cases.json');
const TEMPLATE = path.join(HERE, 'Settings.template.json');
const EVIDENCE = new Set(['accepted', 'rejected']);

const CONTROLS = [
	{
		id: 'control_accept',
		description: 'Control: a valid module with an entry point must build.',
		expected: 'accepted',
		mode: 'compile',
		source: 'Option Explicit\n\nPublic Sub XlideOracleEntry()\n    Dim total As Double\n    total = 100\nEnd Sub\n',
	},
	{
		id: 'control_reject',
		description: 'Control: a trailing comma omitting a required argument must not build.',
		expected: 'rejected',
		mode: 'compile',
		source: 'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n'
			+ '    InvoiceTotal = Subtotal + (Subtotal * TaxRate)\nEnd Function\n\n'
			+ 'Public Sub XlideOracleEntry()\n    Dim total As Double\n    Dim total2 As Double\n    total = 100\n    total2 = InvoiceTotal(total, )\nEnd Sub\n',
	},
];

// ------------------------------------------------------------------ staging

const PROCEDURE = /^[ \t]*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s*\(/gim;
const OPTION_LINE = /^[ \t]*Option\s+(Explicit|Base|Compare|Private)\b.*$/gim;

/** The module's Option lines, and the source without them. */
export function splitOptions(source) {
	const options = [];
	const body = source.replace(OPTION_LINE, (line) => { options.push(line.trim()); return ''; });
	return { options, body };
}

/** The VB6 source of one module as a .twin file, with a touch sub reaching every procedure. */
export function twinModule(name, type, source) {
	const { options, body } = splitOptions(source);
	const procedures = [...body.matchAll(PROCEDURE)].map((m) => m[2]).filter((p) => p.toLowerCase() !== 'main');
	const lines = [];
	if (type === 'class') {
		lines.push(`Class ${name}`, ...options.map((o) => `    ${o}`), body.replace(/\r?\n/g, '\r\n'), 'End Class', '');
		return { text: lines.join('\r\n'), procedures, touch: undefined };
	}
	const touch = `XlideOracleTouch_${name}`;
	lines.push(`Module ${name}`, ...options.map((o) => `    ${o}`), body.replace(/\r?\n/g, '\r\n'));
	// AddressOf forces code generation for a procedure whether or not the
	// case calls it, which is what VBE's Compile does for the whole module.
	lines.push(`    Public Sub ${touch}()`, '        Dim p As LongPtr');
	for (const p of procedures) { lines.push(`        p = AddressOf ${p}`); }
	lines.push('    End Sub', 'End Module', '');
	return { text: lines.join('\r\n'), procedures, touch };
}

/** The case's modules in the corpus' two shapes. */
export function modulesOf(kase) {
	if (Array.isArray(kase.modules) && kase.modules.length) {
		return kase.modules.map((m) => ({ name: String(m.name), type: String(m.type ?? 'standard'), source: String(m.source ?? ''), entry: m.entry === true }));
	}
	return [{ name: 'XlideOracleModule', type: 'standard', source: String(kase.source ?? ''), entry: true }];
}

/**
 * The Excel and Office type libraries as twinBASIC references, read from
 * the registry on this machine: the VBE corpus was verified inside Excel,
 * where both are referenced by default, so a parity run passes --excel to
 * give the 27 cases that name Excel objects the same references. Off by
 * default: the harness must not need Excel for language cases.
 */
export function excelReferences() {
	const libraries = [
		{ id: '{00020813-0000-0000-C000-000000000046}', version: '1.9', name: 'Microsoft Excel 16.0 Object Library', symbolId: 'Excel' },
		{ id: '{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', version: '2.8', name: 'Microsoft Office 16.0 Object Library', symbolId: 'Office' },
	];
	const out = [];
	for (const lib of libraries) {
		const paths = {};
		for (const arch of ['win32', 'win64']) {
			const query = spawnSync('reg', ['query', `HKLM\\SOFTWARE\\Classes\\TypeLib\\${lib.id}\\${lib.version}\\0\\${arch}`, '/ve'], { encoding: 'utf8', timeout: 10_000 });
			const m = (query.stdout ?? '').match(/REG_SZ\s+(.+?)\s*$/m);
			paths[arch] = m ? m[1].trim() : '';
		}
		if (!paths.win32 && !paths.win64) {
			throw new Error(`${lib.name} is not registered on this machine (TypeLib ${lib.id} ${lib.version}); --excel needs Excel installed.`);
		}
		const [major, minor] = lib.version.split('.').map((v) => Number.parseInt(v, 16));
		out.push({ id: lib.id, lcid: 0, name: lib.name, path32: paths.win32, path64: paths.win64, symbolId: lib.symbolId, versionMajor: major, versionMinor: minor });
	}
	return out;
}

export function stageCase(kase, stageRoot, references = []) {
	const dir = fs.mkdtempSync(path.join(stageRoot, `${kase.id.replace(/[^\w-]/g, '_').slice(0, 40)}-`));
	fs.mkdirSync(path.join(dir, 'Sources'));
	const settings = JSON.parse(fs.readFileSync(TEMPLATE, 'utf8'));
	settings['project.references'] = [...settings['project.references'], ...references];
	fs.writeFileSync(path.join(dir, 'Settings'), `${JSON.stringify(settings, null, '\t')}\n`, 'utf8');
	const modules = modulesOf(kase);
	const startup = ['Module XlideOracleStartup', '    Public Sub Main()'];
	const reach = [];
	let definesMain = false;
	for (const mod of modules) {
		const { text, procedures, touch } = twinModule(mod.name, mod.type, mod.source);
		fs.writeFileSync(path.join(dir, 'Sources', `${mod.name}.twin`), text, 'utf8');
		if (procedures.length < [...mod.source.matchAll(PROCEDURE)].length) { definesMain = true; }
		if (mod.type === 'class') {
			// Instantiating a class makes the compiler generate its members.
			startup.push(`        Dim ${mod.name}Instance As New ${mod.name}`, `        Set ${mod.name}Instance = Nothing`);
			reach.push(`${mod.name} (instantiated)`);
		} else if (touch) {
			startup.push(`        ${mod.name}.${touch}`);
			reach.push(`${mod.name} (${procedures.length} procedure${procedures.length === 1 ? '' : 's'} by AddressOf)`);
		}
	}
	const entryModule = modules.find((m) => m.entry && m.type === 'standard') ?? modules.find((m) => m.type === 'standard');
	const entryPoint = kase.entryPoint ?? (modules.some((m) => /\bSub\s+XlideOracleEntry\b/i.test(m.source)) ? 'XlideOracleEntry' : undefined);
	if (entryPoint && entryModule && new RegExp(`\\bSub\\s+${entryPoint}\\b`, 'i').test(entryModule.source)) {
		startup.push(`        ${entryModule.name}.${entryPoint}`);
	}
	startup.push('    End Sub', 'End Module', '');
	fs.writeFileSync(path.join(dir, 'Sources', 'XlideOracleStartup.twin'), startup.join('\r\n'), 'utf8');
	return { dir, project: path.join(dir, 'case.twinproj'), reach, definesMain };
}

// ------------------------------------------------------------------ twinBASIC

function locateTwinBasic(explicit) {
	const candidates = [explicit, process.env.XLIDE_TWINBASIC_DIR].filter(Boolean);
	for (const dir of candidates) {
		if (fs.existsSync(path.join(dir, 'twinBASIC.exe')) && fs.existsSync(path.join(dir, 'bin', 'twinBASIC_win64.exe'))) {
			return path.resolve(dir);
		}
	}
	return undefined;
}

function twinBasicVersion(dir) {
	const completed = spawnSync(path.join(dir, 'bin', 'twinBASIC_win64.exe'), [], { encoding: 'utf8', timeout: 20_000 });
	const banner = `${completed.stdout ?? ''}${completed.stderr ?? ''}`.match(/twinBASIC v[\d.]+/);
	return banner ? banner[0] : 'twinBASIC (version banner not read)';
}

function importProject(dir, stage) {
	const completed = spawnSync(
		path.join(dir, 'bin', 'twinBASIC_win64.exe'),
		['import', stage.project, `${stage.dir}${path.sep}`, '--overwrite'],
		{ encoding: 'utf8', timeout: 60_000 },
	);
	const output = `${completed.stdout ?? ''}${completed.stderr ?? ''}`;
	if (completed.error || !/\.\.\. DONE/.test(output) || !fs.existsSync(stage.project)) {
		return { ok: false, message: (completed.error?.message ?? output.trim().split(/\r?\n/).slice(-3).join(' | ')) || 'import produced no project' };
	}
	return { ok: true };
}

function killTree(pid) {
	spawnSync('taskkill', ['/PID', String(pid), '/T', '/F'], { encoding: 'utf8', timeout: 15_000 });
}

function builtExe(stage) {
	const buildDir = path.join(stage.dir, 'Build');
	if (!fs.existsSync(buildDir)) { return { buildDir: false, exe: undefined }; }
	const exe = fs.readdirSync(buildDir).find((f) => f.toLowerCase().endsWith('.exe'));
	return { buildDir: true, exe: exe ? path.join(buildDir, exe) : undefined };
}

const sleep = (ms) => new Promise((resolve) => { setTimeout(resolve, ms); });

/**
 * Builds one staged case. Resolves to { outcome, stage, message, seconds }.
 *
 * The clock that matters starts when `Build\` appears: a successful build
 * writes its EXE and exits within a second of creating the folder
 * (measured), so a folder that stays empty for `stallSeconds` while the IDE
 * stays up is the compiler having refused the build.
 */
async function build(dir, stage, timeoutSeconds, stallSeconds, startupSeconds = timeoutSeconds) {
	const child = spawn(path.join(dir, 'twinBASIC.exe'), [stage.project, '--buildAndExit64'], {
		cwd: dir, stdio: 'ignore', windowsHide: false,
	});
	const started = Date.now();
	let exited = false;
	let exitCode;
	let buildDirSeen;
	child.on('exit', (code) => { exited = true; exitCode = code; });
	child.on('error', () => { exited = true; exitCode = -1; });
	const elapsed = () => (Date.now() - started) / 1000;
	for (;;) {
		await sleep(200);
		if (exited) {
			const { exe } = builtExe(stage);
			if (exe) { return { outcome: 'accepted', stage: 'build', message: `built ${path.basename(exe)}`, seconds: elapsed() }; }
			return { outcome: 'build_incomplete', stage: 'build', message: `the IDE exited (${exitCode}) without writing an EXE`, seconds: elapsed() };
		}
		const { buildDir, exe } = builtExe(stage);
		if (buildDir && buildDirSeen === undefined) { buildDirSeen = Date.now(); }
		const sinceBuildDir = buildDirSeen === undefined ? 0 : (Date.now() - buildDirSeen) / 1000;
		if (exe && sinceBuildDir > stallSeconds) {
			// Built but still up: not the documented stall; kill and report it.
			killTree(child.pid);
			return { outcome: 'accepted', stage: 'build', message: `built ${path.basename(exe)}; the IDE stayed up and was killed`, seconds: elapsed() };
		}
		if (buildDir && !exe && sinceBuildDir >= stallSeconds) {
			killTree(child.pid);
			return {
				outcome: 'rejected', stage: 'build_failed',
				message: `the compiler reported errors: the IDE stalled with an empty Build folder for ${stallSeconds}s (its documented behaviour)`,
				seconds: elapsed(),
			};
		}
		if (!buildDir && elapsed() >= startupSeconds) {
			// An IDE that has not started the build by now is stuck starting
			// up (a successful build creates Build\ within seconds), not
			// compiling; the caller retries once.
			killTree(child.pid);
			return { outcome: 'timeout', stage: 'startup', message: `the IDE did not start the build within ${startupSeconds}s`, seconds: elapsed() };
		}
		if (elapsed() >= timeoutSeconds) {
			killTree(child.pid);
			return { outcome: 'timeout', stage: buildDir ? 'build' : 'startup', message: `no verdict after ${timeoutSeconds}s (Build folder ${buildDir ? 'present' : 'absent'})`, seconds: elapsed() };
		}
	}
}

/**
 * Removes a staged folder. The IDE's children (compiler, WebView2) release
 * their handles a moment after the IDE itself exits, so the first attempt
 * can hit EPERM; a folder that stays busy is left behind and reported, never
 * allowed to abort the batch.
 */
function removeStaged(dir) {
	try {
		fs.rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 500 });
		return undefined;
	} catch (err) {
		return `staged folder left in place (${err.code ?? err.message}): ${dir}`;
	}
}

async function runCase(kase, dir, options) {
	const stage = stageCase(kase, options.stageRoot, options.references);
	let result;
	try {
		if (stage.definesMain) {
			result = { caseId: kase.id, outcome: 'unsupported', stage: 'setup', message: 'the case declares Sub Main, which the harness reserves for its startup module' };
		} else {
			const imported = importProject(dir, stage);
			if (!imported.ok) {
				result = { caseId: kase.id, outcome: 'worker_error', stage: 'import', message: imported.message };
			} else {
				let built = await build(dir, stage, options.timeout, options.stall, options.startup);
				let attempts = 1;
				if (built.outcome === 'timeout' && built.stage === 'startup') {
					// A startup hang is the instance, not the case: one retry.
					attempts += 1;
					built = await build(dir, stage, options.timeout, options.stall, options.startup);
				}
				result = { caseId: kase.id, reach: stage.reach, ...built, ...(attempts > 1 ? { attempts } : {}) };
			}
		}
	} finally {
		if (!options.keep) {
			const leftover = removeStaged(stage.dir);
			if (leftover && result) { result.cleanup = leftover; }
		}
	}
	return result;
}

// ------------------------------------------------------------------ evidence

export // Cases that name Excel's own objects: their verdict rides on the Excel type
// library the project references, and a disagreement there is about the
// reference, not the language. Tagged so the parity doc can say which is which.
const EXCEL_HOST_NAMES = /\b(Range|Cells|Worksheets?|Workbooks?|ThisWorkbook|ActiveSheet|ActiveWorkbook|ActiveCell|Application|Sheets|Charts?|Selection|WorksheetFunction|Names|Union|Intersect|ListObjects?|Shapes?)\b/;

export function hostBound(kase) {
	return modulesOf(kase).some((m) => EXCEL_HOST_NAMES.test(m.source));
}

export function compileExpectation(kase) {
	// A VBE runtime case compiled before it ran, so at compile it is accepted.
	const phase = String(kase.evidencePhase ?? '');
	if (phase === 'runtime' || String(kase.mode ?? '') === 'run') { return 'accepted'; }
	return String(kase.expected ?? 'observe');
}

function matches(expected, outcome) {
	if (!EVIDENCE.has(outcome)) { return false; }
	return expected === 'observe' || expected === outcome;
}

export function parityDoc(results, controls, header) {
	const rows = results.filter((r) => EVIDENCE.has(r.outcome) && r.compileExpected !== 'observe');
	const cell = (e, o) => rows.filter((r) => r.compileExpected === e && r.outcome === o).length;
	const lines = [];
	lines.push('# twinBASIC parity with the Excel/VBE oracle corpus');
	lines.push('');
	lines.push(`Generated by \`syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs\` on ${header.date}`);
	lines.push(`against ${header.version} (${header.dir})${header.excel ? ', with the Excel and Office type libraries referenced' : ', with no Excel reference'};`);
	lines.push('regenerate rather than edit.');
	lines.push('');
	lines.push('Each VBE-verified case was staged as a twinBASIC project and built headlessly.');
	lines.push('The expectation compared is the case at COMPILE time: a VBE runtime case');
	lines.push('compiled before it ran, so it counts as accepted here. twinBASIC is a');
	lines.push('superset of VB6 with published incompatibilities: agreement is evidence that');
	lines.push('the oracle can stand in for VB6 on that case, disagreement is a recorded');
	lines.push('difference, and neither is a VB6 verdict.');
	lines.push('');
	lines.push('## Matrix');
	lines.push('');
	lines.push('| VBE expected (compile) | twinBASIC accepted | twinBASIC rejected |');
	lines.push('| --- | --- | --- |');
	lines.push(`| accepted | ${cell('accepted', 'accepted')} | ${cell('accepted', 'rejected')} |`);
	lines.push(`| rejected | ${cell('rejected', 'accepted')} | ${cell('rejected', 'rejected')} |`);
	lines.push('');
	const agree = cell('accepted', 'accepted') + cell('rejected', 'rejected');
	lines.push(`Agreement: ${agree} of ${rows.length} (${rows.length ? Math.round((agree / rows.length) * 1000) / 10 : 0}%).`);
	const language = rows.filter((r) => !r.hostBound);
	const languageAgree = language.filter((r) => r.compileExpected === r.outcome).length;
	lines.push(`Language-only cases (naming no Excel object): ${languageAgree} of ${language.length} agree`
		+ ` (${language.length ? Math.round((languageAgree / language.length) * 1000) / 10 : 0}%);`
		+ ` Excel-bound cases: ${agree - languageAgree} of ${rows.length - language.length} agree.`);
	const other = results.filter((r) => !EVIDENCE.has(r.outcome));
	lines.push(`Not evidence: ${other.length} case(s) (${[...new Set(other.map((r) => r.outcome))].join(', ') || 'none'}).`);
	lines.push('');
	lines.push('## Controls');
	lines.push('');
	for (const c of controls) { lines.push(`- ${c.when} ${c.caseId}: ${c.outcome} (expected ${c.expected})`); }
	lines.push('');
	lines.push('## Disagreements');
	lines.push('');
	const disagreements = rows.filter((r) => r.compileExpected !== r.outcome);
	if (!disagreements.length) { lines.push('- none'); }
	for (const r of disagreements) {
		lines.push(`- ${r.caseId}${r.hostBound ? ' [Excel-bound]' : ''}: VBE ${r.compileExpected}, twinBASIC ${r.outcome} (${r.stage}). ${r.description ?? ''}`.trimEnd());
	}
	lines.push('');
	lines.push('## Not evidence');
	lines.push('');
	if (!other.length) { lines.push('- none'); }
	for (const r of other) { lines.push(`- ${r.caseId}: ${r.outcome} (${r.stage}) ${r.message ?? ''}`.trimEnd()); }
	lines.push('');
	return lines.join('\n');
}

// ------------------------------------------------------------------ main

function parseArgs(argv) {
	const args = { cases: DEFAULT_CASES, caseIds: [], limit: Infinity, offset: 0, append: false, twinbasic: undefined, timeout: 60, stall: 4, startup: 20, parallel: 1, stagger: 1500, excel: false, controls: true, keep: false, json: false, report: undefined, parityDoc: undefined, strict: false };
	const number = (raw, flag) => { const v = Number.parseInt(raw, 10); if (!Number.isInteger(v)) { throw new Error(`${flag} expects an integer`); } return v; };
	for (let i = 0; i < argv.length; i += 1) {
		const flag = argv[i];
		switch (flag) {
			case '--cases': args.cases = argv[++i]; break;
			case '--case': args.caseIds.push(argv[++i]); break;
			case '--limit': args.limit = number(argv[++i], flag); break;
			case '--offset': args.offset = number(argv[++i], flag); break;
			case '--append': args.append = true; break;
			case '--twinbasic': args.twinbasic = argv[++i]; break;
			case '--timeout': args.timeout = number(argv[++i], flag); break;
			case '--stall': args.stall = number(argv[++i], flag); break;
			case '--parallel': args.parallel = Math.max(1, number(argv[++i], flag)); break;
			case '--startup': args.startup = number(argv[++i], flag); break;
			case '--stagger': args.stagger = number(argv[++i], flag); break;
			case '--excel': args.excel = true; break;
			case '--no-controls': args.controls = false; break;
			case '--keep': args.keep = true; break;
			case '--json': args.json = true; break;
			case '--report': args.report = argv[++i]; break;
			case '--parity-doc': args.parityDoc = argv[++i]; break;
			case '--strict': args.strict = true; break;
			case '-h': case '--help': args.help = true; break;
			default: throw new Error(`Unknown argument: ${flag}`);
		}
	}
	return args;
}

async function main(argv) {
	let args;
	try { args = parseArgs(argv); } catch (err) { console.error(err.message); return 2; }
	if (args.help) { console.log(fs.readFileSync(fileURLToPath(import.meta.url), 'utf8').split('\n').filter((l) => l.startsWith('//')).map((l) => l.slice(3)).join('\n')); return 0; }
	if (process.platform !== 'win32') { console.error('The twinBASIC oracle requires Windows.'); return 2; }
	const dir = locateTwinBasic(args.twinbasic);
	if (!dir) {
		console.error('twinBASIC not found: pass --twinbasic <dir> or set XLIDE_TWINBASIC_DIR to the IDE folder (twinBASIC.exe, bin\\twinBASIC_win64.exe). Nothing is bundled.');
		return 2;
	}
	const version = twinBasicVersion(dir);
	const document = JSON.parse(fs.readFileSync(args.cases, 'utf8').replace(/^﻿/, ''));
	let cases = document.cases;
	if (!Array.isArray(cases)) { console.error(`${args.cases} has no cases array`); return 2; }
	if (args.caseIds.length) {
		const wanted = new Set(args.caseIds);
		cases = cases.filter((k) => wanted.has(k.id));
		const missing = args.caseIds.filter((id) => !cases.some((k) => k.id === id));
		if (missing.length) { console.error(`Unknown case(s): ${missing.join(', ')}`); return 2; }
	}
	// Earlier evidence in the report is kept and its cases skipped, so a
	// long batch can be run in chunks; controls still run for every chunk.
	let earlier = { controls: [], results: [] };
	if (args.append && args.report && fs.existsSync(args.report)) {
		earlier = JSON.parse(fs.readFileSync(args.report, 'utf8'));
		const done = new Set((earlier.results ?? []).filter((r) => EVIDENCE.has(r.outcome) || r.outcome === 'unsupported').map((r) => r.caseId));
		cases = cases.filter((k) => !done.has(k.id));
	}
	cases = cases.slice(args.offset, args.offset + args.limit);
	let references = [];
	if (args.excel) {
		try { references = excelReferences(); } catch (err) { console.error(err.message); return 2; }
	}
	const options = { timeout: args.timeout, stall: args.stall, startup: args.startup, keep: args.keep, references, stageRoot: fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-twinbasic-oracle-')) };

	const log = (line) => { if (!args.json) { console.log(line); } };
	const controls = [];
	async function runControls(when) {
		if (!args.controls) { return true; }
		let ok = true;
		for (const control of CONTROLS) {
			const result = await runCase(control, dir, options);
			controls.push({ when, caseId: control.id, expected: control.expected, outcome: result.outcome, stage: result.stage, message: result.message });
			const pass = result.outcome === control.expected;
			log(`${pass ? 'CONTROL-PASS' : 'CONTROL-FAIL'} ${when} ${control.id}: outcome=${result.outcome} expected=${control.expected} (${Math.round(result.seconds ?? 0)}s)`);
			ok = ok && pass;
		}
		return ok;
	}

	const results = [];
	let failures = 0;
	let oracleFailures = 0;
	const header = { date: new Date().toISOString().slice(0, 10), version, dir: path.basename(dir), excel: args.excel };
	// The report is rewritten after every case, so a crash or a kill keeps
	// the evidence gathered so far and --append can resume from it.
	function writeReport() {
		if (!args.report) { return; }
		const allControls = [...(earlier.controls ?? []), ...controls];
		const allResults = [...(earlier.results ?? []).filter((r) => !results.some((n) => n.caseId === r.caseId)), ...results];
		fs.writeFileSync(args.report, `${JSON.stringify({ generatedAt: header.date, twinbasic: { version, dir: header.dir }, cases: path.basename(args.cases), excelReferenced: args.excel, controls: allControls, results: allResults }, null, 2)}\n`, 'utf8');
		return allResults;
	}
	const controlsBefore = await runControls('before');
	if (!controlsBefore) {
		console.error('ORACLE-FAIL: a control did not answer as expected; the harness, not the cases, needs attention.');
		oracleFailures += 1;
	} else {
		// A small pool: each case has its own folder and IDE process tree,
		// so instances do not share anything but the machine. Results are
		// recorded in corpus order whatever the finish order.
		let next = 0;
		let stop = false;
		const record = (kase, result) => {
			result.expected = String(kase.expected ?? 'observe');
			result.compileExpected = compileExpectation(kase);
			result.matched = matches(result.compileExpected, result.outcome);
			result.description = kase.description ?? '';
			result.evidencePhase = 'compile';
			result.hostBound = hostBound(kase);
			results.push(result);
			const infra = !EVIDENCE.has(result.outcome) && result.outcome !== 'unsupported';
			if (infra) { oracleFailures += 1; } else if (result.outcome !== 'unsupported' && !result.matched) { failures += 1; }
			const marker = infra ? 'ORACLE-FAIL' : (result.outcome === 'unsupported' ? 'SKIP' : (result.matched ? 'PASS' : 'FAIL'));
			log(`${marker} ${kase.id}: outcome=${result.outcome} expected=${result.compileExpected} stage=${result.stage} (${Math.round(result.seconds ?? 0)}s)`);
			if (result.message && (infra || !result.matched)) { log(`  ${result.message}`); }
			if (result.cleanup) { log(`  ${result.cleanup}`); }
			writeReport();
			// A timeout inside the build (Build\ present, no verdict) is the
			// systemic kind and stops the batch; a startup hang that survived
			// its retry is instance-local and is recorded, the batch goes on.
			if (infra && result.outcome === 'timeout' && result.stage === 'build') { stop = true; }
		};
		// Instances start staggered: several IDEs launched in the same
		// instant leave one stuck before it builds (measured with four).
		const worker = async (index) => {
			await sleep(index * args.stagger);
			while (!stop && next < cases.length) {
				const kase = cases[next++];
				record(kase, await runCase(kase, dir, options));
			}
		};
		await Promise.all(Array.from({ length: Math.min(args.parallel, cases.length || 1) }, (_, index) => worker(index)));
		results.sort((a, b) => cases.findIndex((k) => k.id === a.caseId) - cases.findIndex((k) => k.id === b.caseId));
		const controlsAfter = await runControls('after');
		if (!controlsAfter) {
			console.error('ORACLE-FAIL: a control failed after the batch; this batch\'s rejections are not evidence.');
			oracleFailures += 1;
			for (const r of results) { if (r.outcome === 'rejected') { r.outcome = 'unverified'; r.matched = false; } }
		}
	}
	if (!args.keep) {
		const leftover = removeStaged(options.stageRoot);
		if (leftover) { log(leftover); }
	}

	const allControls = [...(earlier.controls ?? []), ...controls];
	const allResults = writeReport() ?? results;
	if (args.report) { log(`Wrote ${args.report} (${allResults.length} result(s))`); }
	if (args.parityDoc) {
		fs.writeFileSync(args.parityDoc, parityDoc(allResults, allControls, header), 'utf8');
		log(`Wrote ${args.parityDoc}`);
	}
	if (args.json) {
		console.log(JSON.stringify({ twinbasic: { version, dir: header.dir }, controls, results, failureCount: failures, oracleFailureCount: oracleFailures }, null, 2));
	} else {
		log(`\n${results.length} case(s), ${failures} expectation mismatch(es), ${oracleFailures} oracle infrastructure failure(s). ${version}.`);
	}
	return oracleFailures || (args.strict && failures) ? 1 : 0;
}

const invokedDirectly = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) {
	process.exit(await main(process.argv.slice(2)));
}
