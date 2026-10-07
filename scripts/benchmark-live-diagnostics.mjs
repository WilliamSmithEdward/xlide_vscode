// Read-only live-edit benchmark: retains worker state across changed snapshots.
// node scripts/benchmark-live-diagnostics.mjs <workbook.xlsm> [--profile]
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { mkdtempSync, rmSync, writeFileSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';
import { Session } from 'node:inspector';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const input = process.argv.slice(2).find(arg => !arg.startsWith('--'));
if (!input) throw new Error('Pass a workbook path. The workbook is only read.');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-live-diagnostics-'));
try {
    const bundle = join(scratch, 'api.cjs');
    const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice('--baseline='.length);
    const plugins = baseline ? [{ name: 'saved-live-diagnostics', setup(build) {
        build.onLoad({ filter: /(?:analysisWorkerLogic|deadCode)\.ts$/ }, args => {
            const relative = args.path.slice(root.length + 1).replaceAll('\\', '/');
            if (!['src/analysisWorkerLogic.ts', 'src/analyzer/diagnostics/rules/deadCode.ts'].includes(relative)) return;
            return { contents: execFileSync('git', ['show', `${baseline}:${relative}`], { cwd: root, encoding: 'utf8' }), loader: 'ts' };
        });
    } }] : [];
    if (process.argv.includes('--dirty')) plugins.push({ name: 'dirty-procedures', setup(build) {
        build.onLoad({ filter: /diagnostics[\\/]incrementalRules\.ts$/ }, args => ({
            contents: readFileSync(args.path, 'utf8').replace('const dirtyStarts = new Set<number>',
                'console.log(JSON.stringify({ dirty: [...dirty].map(i => members[i].name) })); const dirtyStarts = new Set<number>'), loader: 'ts',
        }));
    } });
    // Controlled comparison: keep every analyzer change, but restore the two
    // per-procedure Map copies and ordinary mutable-start detachment.
    if (process.argv.includes('--copied-starts')) plugins.push({ name: 'copied-starts', setup(build) {
        build.onLoad({ filter: /diagnostics[\\/]typeInference\.ts$/ }, args => {
            const contents = readFileSync(args.path, 'utf8');
            const shared = 'start = shareImmutableReachingStart(conditionConstants(symbols, proc, new Map([...declaredDefaults(locals), ...variantStarts(symbols, proc), ...objectStarts(symbols, proc)])));';
            if (!contents.includes(shared)) throw new Error('Cannot find shared procedure start for comparison');
            return { contents: contents.replace(shared, 'start = new Map([...new Map(conditionConstants(symbols, proc, new Map())), ...declaredDefaults(locals), ...variantStarts(symbols, proc), ...objectStarts(symbols, proc)]);'), loader: 'ts' };
        });
    } });
    await build({ stdin: { contents: `
        export { readModules } from './src/vba/projectService';
        export { AnalysisWorkerState } from './src/analysisWorkerLogic';
        export { moduleKindFromType } from './src/vbaProjectAnalysis';
        export { parseModule } from './src/analyzer/parser/parseModule';
        export { createConditionalActivityTracker } from './src/analyzer/conditional/conditionalCompilation';
        export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';
        export { DIAGNOSTIC_RULE_REGISTRY } from './src/analyzer/diagnostics/registry';
        export { checkUnusedDeclarations } from './src/analyzer/diagnostics/rules/deadCode';
        export { hostTokenForFileName } from './src/analyzer/host/hostRegistry';
        export { referencedHostTokens } from './src/analyzer/host/hostLibraries';
    `, resolveDir: root, loader: 'ts' }, outfile: bundle, bundle: true, platform: 'node', format: 'cjs', sourcemap: true, plugins });
    const api = createRequire(import.meta.url)(bundle);
    const workbook = resolve(input);
    const modules = api.readModules(workbook);
    const mod = [...modules].sort((a, b) => b.source.length - a.source.length)[0];
    if (process.argv.includes('--unused-only')) {
        const parsed = api.parseModule(mod.source);
        const symbols = api.buildModuleSymbols(mod.name, api.moduleKindFromType(mod.type), mod.source);
        const activity = api.createConditionalActivityTracker(parsed);
        const samples = [];
        let expected;
        for (let i = 0; i < 10; i++) {
            const findings = [];
            const start = performance.now();
            api.checkUnusedDeclarations(mod.source, parsed, symbols, activity, (...args) => findings.push(args));
            const ms = performance.now() - start;
            const digest = createHash('sha256').update(JSON.stringify(findings)).digest('hex');
            expected ??= digest;
            if (expected !== digest) throw new Error('Rule diagnostics changed between rounds');
            if (i >= 3) samples.push(ms);
        }
        console.log(JSON.stringify({ module: mod.name, baseline, medianMs: samples.sort((a,b) => a-b)[3], diagnosticDigest: expected }));
    } else {
    const parsed = api.parseModule(mod.source);
    const activity = api.createConditionalActivityTracker(parsed);
    const closures = process.argv.includes('--closures');
    const screenshot = process.argv.includes('--screenshot') || closures;
    const procedure = parsed.members.find(m => m.kind === 'Procedure' && !activity?.isInactive(m.span)
        && (screenshot ? m.name === 'Value' : m.body.length > 2));
    if (!procedure) throw new Error('No nonempty procedure to edit');
    // Insert inside an existing body, preserving the declaration/signature envelope.
    const at = screenshot ? mod.source.lastIndexOf('End Function', procedure.span.end - 1) : procedure.body[1].span.start;
    const worker = new api.AnalysisWorkerState();
    worker.handle({ kind: 'seed', projectKey: 'live', generation: 1,
        modules: modules.map(m => ({ ...m, moduleName: m.name })) });
    const host = api.hostTokenForFileName(workbook);
    const references = modules[0]?.projectReferences ?? [];
    const base = { kind: 'analyze', requestId: 1, docKey: mod.name, projectKey: 'live', generation: 1,
        ...(process.argv.includes('--cancellable') ? { cancellationSignal: new Int32Array(new SharedArrayBuffer(4)) } : {}),
        errorsOnly: process.argv.includes('--early'),
        moduleName: mod.name, moduleType: mod.type, moduleKind: api.moduleKindFromType(mod.type),
        documentType: mod.documentType, host, referencedHosts: api.referencedHostTokens(host, references),
        referencedLibraries: references.length ? references.map(r => r.name) : undefined,
        designerClass: mod.designerClass, workbookSheets: modules[0]?.projectSheets };
    const ruleTimes = new Map();
    if (process.argv.includes('--rules')) {
        for (const rule of api.DIAGNOSTIC_RULE_REGISTRY) {
            if (!rule.run) continue;
            const runRule = rule.run;
            rule.run = (...args) => { const start = performance.now(); try { return runRule(...args); } finally { ruleTimes.set(rule.name, (ruleTimes.get(rule.name) ?? 0) + performance.now() - start); } };
        }
    }
    const rows = [];
    const digest = createHash('sha256');
    const canonical = result => JSON.stringify([result.diagnostics, result.suppressedDiagnostics].map(list => list.map(d => {
        const { origin, walkMemberStart, ...publicDiagnostic } = d;
        return JSON.stringify(publicDiagnostic);
    }).sort()));
    function run(label, source) {
        const previous = worker._incrementalByDoc.get(mod.name);
        ruleTimes.clear();
        const start = performance.now();
        const result = worker.handle({ ...base, source });
        const ms = performance.now() - start;
        if (result?.kind !== 'result' || result.analysisFailures?.length) throw new Error(JSON.stringify(result));
        digest.update(JSON.stringify([label, result.diagnostics, result.suppressedDiagnostics]));
        const next = worker._incrementalByDoc.get(mod.name);
        const oldRecords = new Map(previous?.procedures.map(p => [p.key, p]) ?? []);
        rows.push({ label, ms: +ms.toFixed(2), mode: result.incrementalMode,
            ...(ruleTimes.size ? { rules: [...ruleTimes].sort((a,b)=>b[1]-a[1]).slice(0,8) } : {}),
            ...(previous && next ? { envelopeChanged: previous.envelope !== next.envelope,
                effectChanges: next.procedures.filter(p => p.effectText !== oldRecords.get(p.key)?.effectText).length } : {}),
            ...(screenshot && next ? { targetReferences: next.procedures.filter(p => p.references?.has(procedure.name.toLowerCase())).length,
                procedures: next.procedures.length } : {}),
            diagnostics: result.diagnostics.length,
            editErrors: result.diagnostics.filter(d => d.severity === 'error' && d.span.start >= at && d.span.start < at + (screenshot ? 120 : 35)).map(d => d.code) });
        if (process.argv.includes('--verify')) {
            const oracle = new api.AnalysisWorkerState();
            oracle.handle({ kind: 'seed', projectKey: 'live', generation: 1, modules: modules.map(m => ({ ...m, moduleName: m.name })) });
            const full = oracle.handle({ ...base, cancellationSignal: undefined, source });
            if (full?.kind !== 'result' || full.analysisFailures?.length || canonical(result) !== canonical(full)) {
                throw new Error(JSON.stringify({ label, incremental: result, full }));
            }
            rows.at(-1).fullParity = true;
        }
    }
    run('initial', mod.source);
    ruleTimes.clear();
    const session = new Session();
    const post = (method, params = {}) => new Promise((resolve, reject) => session.post(method, params, (err, result) => err ? reject(err) : resolve(result)));
    if (process.argv.includes('--profile')) { session.connect(); await post('Profiler.enable'); if (!screenshot) await post('Profiler.start'); }
    if (closures) {
        if (process.argv.includes('--profile')) await post('Profiler.start');
        for (let i = 0; i < 3; i++) {
            for (const closer of ['End Sub', '']) {
                run(closer ? 'mismatched closer' : 'missing closer', mod.source.slice(0, at) + closer + mod.source.slice(at + 'End Function'.length));
                run('closer repair', mod.source);
            }
        }
    } else if (screenshot) {
        const snapshot = statement => mod.source.slice(0, at) + statement + mod.source.slice(at);
        run('screenshot warmup', snapshot('    Dim wb As Workbook\n    Set wb = ThisWorkbook\n'));
        if (process.argv.includes('--profile')) await post('Profiler.start');
        for (let i = 0; i < 3; i++) {
            run('screenshot declaration typo', snapshot('    Debug wb As Worksheet\n    Set wb = ThisWorkbook\n'));
            run('screenshot declaration repair', snapshot('    Dim wb As Worksheet\n    Set wb = ThisWorkbook\n'));
            run('screenshot assignment repair', snapshot('    Dim wb As Workbook\n    Set wb = ThisWorkbook\n'));
        }
    } else if (process.argv.includes('--surface')) {
        const beforeProcedure = procedure.span.start;
        for (let i = 0; i < 3; i++) {
            run('module declaration addition', mod.source.slice(0, beforeProcedure)
                + `Private Const XlidePerfConstant As Long = ${i}\n` + mod.source.slice(beforeProcedure));
            run('module declaration removal', mod.source);
            run('procedure addition', mod.source + '\nPrivate Sub XlidePerfNewProcedure()\nDim unused As Long\nEnd Sub\n');
            run('procedure removal', mod.source);
            run('callee effect edit', mod.source.slice(0, at) + `Dim XlidePerfLocal As Long\nXlidePerfLocal = ${i}\n` + mod.source.slice(at));
            run('callee effect removal', mod.source);
        }
    } else for (let i = 0; i < 5; i++) {
        run('valid edit', mod.source.slice(0, at) + `Debug.Print ${i}\n` + mod.source.slice(at));
        const edit = process.argv.find(arg => arg.startsWith('--statement='))?.slice('--statement='.length) ?? 'Debug.Print (1';
        const broken = mod.source.slice(0, at) + edit + '\n' + mod.source.slice(at);
        run('hard syntax error', broken);
        if (!rows.at(-1).editErrors.length) throw new Error('Expected hard error at the edit');
        run('unchanged follow-up', broken);
    }
    if (process.argv.includes('--profile')) {
        const { profile } = await post('Profiler.stop'); session.disconnect();
        writeFileSync(join(root, 'live-diagnostics.cpuprofile'), JSON.stringify(profile));
        const counts = new Map();
        for (const id of profile.samples ?? []) counts.set(id, (counts.get(id) ?? 0) + 1);
        const hot = profile.nodes.map(n => ({ name: n.callFrame.functionName, line: n.callFrame.lineNumber + 1, samples: counts.get(n.id) ?? 0 })).sort((a,b) => b.samples-a.samples).slice(0,25);
        console.log(JSON.stringify({ hot }, null, 2));
    }
    if (ruleTimes.size) console.log(JSON.stringify({ ruleTimes: [...ruleTimes].sort((a,b)=>b[1]-a[1]).slice(0,25) }));
    const summary = [...new Set(rows.map(row => row.label))].map(label => {
        const samples = rows.filter(row => row.label === label).map(row => row.ms).sort((a,b) => a-b);
        return { label, medianMs: samples[Math.floor(samples.length / 2)], minMs: samples[0], maxMs: samples.at(-1) };
    });
    console.log(JSON.stringify({ module: mod.name, bytes: mod.source.length, editProcedure: procedure.name,
        baseline, copiedStarts: process.argv.includes('--copied-starts'), diagnosticDigest: digest.digest('hex'),
        ...(process.argv.includes('--verify') ? { verifiedSnapshots: rows.filter(row => row.fullParity).length } : {}),
        ...(process.argv.includes('--summary') ? { summary } : { rows }) }, null, 2));
    }
} finally { rmSync(scratch, { recursive: true, force: true }); }
