// Run: node scripts/benchmark-dim-initializer-context.mjs [--baseline=COMMIT]
import {build} from 'esbuild';
import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';
import {mkdtempSync, writeFileSync, unlinkSync, rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {cpus, tmpdir} from 'node:os';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
const root = dirname(dirname(fileURLToPath(import.meta.url))), baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-dim-context-')), file = join(scratch, 'api.cjs');
let api;
try {
	const plugins = baseline ? [{name: 'baseline', setup(builder) {
		builder.onLoad({filter: /[\\/]codeActions[\\/]diagnosticCodeActions\.ts$/}, args => ({contents: execFileSync('git', ['show', baseline + ':src/analyzer/codeActions/diagnosticCodeActions.ts'], {cwd: root, encoding: 'utf8'}), loader: 'ts', resolveDir: dirname(args.path)}));
	}}] : [];
	const result = await build({plugins, stdin: {contents: "export {resolveDiagnosticCodeActions} from './src/analyzer/codeActions/diagnosticCodeActions';", resolveDir: root, loader: 'ts'}, bundle: true, platform: 'node', format: 'cjs', write: false});
	writeFileSync(file, result.outputFiles[0].contents);
	api = createRequire(import.meta.url)(file);
} finally {
	try { unlinkSync(file); } catch (error) { if (error.code !== 'ENOENT') { throw error; } }
	rmdirSync(scratch);
}
const rows = [];
function measure(label, source, inside) {
	const offset = source.indexOf('='), end = source.indexOf('\n', offset), diagnostic = {code: 'dim-initializer', span: {start: offset, end: offset + 1}};
	const expected = inside ? [{title: 'Split declaration initializer', kind: 'quickfix', isPreferred: true, edits: [{span: {start: offset - 1, end: end < 0 ? source.length : end}, newText: '\n  value = 42'}]}] : [];
	const samples = [];
	for (let round = -3; round < 9; round++) {
		let actual;
		const start = performance.now();
		for (let i = 0; i < 20; i++) { actual = api.resolveDiagnosticCodeActions(source, diagnostic); }
		const elapsed = (performance.now() - start) / 20;
		assert.deepEqual(actual, expected);
		if (round >= 0) { samples.push(elapsed); }
	}
	samples.sort((a, b) => a - b);
	rows.push({label, sourceCharacters: source.length, medianMs: samples[4], maxBatchAverageMs: samples[8]});
}
for (const count of [1, 100, 10000]) {
	const unrelated = Array.from({length: count}, (_, i) => 'Sub P' + i + '()\nDim n As Long\nEnd Sub\n');
	for (const [position, before] of [['first', 0], ['middle', Math.floor(count / 2)], ['last', count]]) {
		measure(count + '-' + position, unrelated.slice(0, before).join('') + 'Sub Target()\n  Dim value As Long = 42\nEnd Sub\n' + unrelated.slice(before).join(''), true);
	}
}
measure('long-first-body', 'Sub Target()\n' + 'Dim n As Long\n'.repeat(30000) + '  Dim value As Long = 42\nEnd Sub\n', true);
measure('module-level-after-10000-procedures', Array.from({length: 10000}, (_, i) => 'Sub P' + i + '()\nDim n As Long\nEnd Sub\n').join('') + '  Dim value As Long = 42\n', false);
console.log(JSON.stringify({baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, batch: 20,
	scope: 'Complete resolveDiagnosticCodeActions query, including procedure context, tokenization, EOL detection and edits. Input and independent complete action assertions outside timer. Statement lexer warmed. No cold-lexer/complete-editor latency claim.', rows}, null, 2));
