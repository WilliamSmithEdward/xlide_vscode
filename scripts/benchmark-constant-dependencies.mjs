// Run: node scripts/benchmark-constant-dependencies.mjs [--baseline=COMMIT]
import assert from 'node:assert/strict';
import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync, writeFileSync, unlinkSync, rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {cpus, tmpdir} from 'node:os';
import {dirname, join} from 'node:path';
import {performance} from 'node:perf_hooks';
import {fileURLToPath} from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-constant-dependencies-'));
const file = join(scratch, 'api.cjs');
let api;
try {
	const plugins = baseline ? [{name: 'baseline', setup(builder) {
		builder.onLoad({filter: /[\\/]constants[\\/]integerConstantExpression\.ts$/}, args => ({
			contents: execFileSync('git', ['show', baseline + ':src/analyzer/constants/integerConstantExpression.ts'], {cwd: root, encoding: 'utf8'}),
			loader: 'ts', resolveDir: dirname(args.path),
		}));
	}}] : [];
	const result = await build({plugins, stdin: {
		contents: "export {resolveRawIntegerConstants, evaluateIntegerConstantExpression} from './src/analyzer/constants/integerConstantExpression';",
		resolveDir: root, loader: 'ts',
	}, bundle: true, platform: 'node', format: 'cjs', write: false});
	writeFileSync(file, result.outputFiles[0].contents);
	api = createRequire(import.meta.url)(file);
} finally {
	try { unlinkSync(file); } catch (error) { if (error.code !== 'ENOENT') { throw error; } }
	rmdirSync(scratch);
}

const rows = [];
function measure(label, run, verify, batch, allowOverflow = false) {
	const samples = [];
	let overflowSamples = 0;
	for (let round = -3; round < 9; round++) {
		let result, error;
		const start = performance.now();
		for (let i = 0; i < batch; i++) {
			try { result = run(); } catch (caught) {
				if (!allowOverflow || !(caught instanceof RangeError)) { throw caught; }
				error = caught;
			}
		}
		const elapsed = (performance.now() - start) / batch;
		if (!error) { verify(result); }
		if (round >= 0) {
			samples.push(elapsed);
			if (error) { overflowSamples++; }
		}
	}
	samples.sort((a, b) => a - b);
	rows.push({label, batch, medianMs: samples[4], maxBatchAverageMs: samples[8], overflowSamples});
}

for (const count of [1, 100, 10000]) {
	for (const mode of ['forward', 'reverse', 'cycle', 'fanout']) {
		let entries = Array.from({length: count}, (_, i) => ['c' + i,
			i === count - 1 ? (mode === 'cycle' ? 'c0' : '1') : 'c' + (i + 1) + ' + 1']);
		let expected = Array.from({length: count}, (_, i) => ['c' + (count - 1 - i), mode === 'cycle' ? undefined : i + 1]);
		if (mode === 'reverse') { entries.reverse(); }
		if (mode === 'cycle') { expected.unshift(['c0', undefined]); expected.pop(); }
		if (mode === 'fanout') {
			entries = [['root', Array.from({length: count}, (_, i) => 'c' + i).join(' + ')],
				...Array.from({length: count}, (_, i) => ['c' + i, '1'])];
			expected = [...Array.from({length: count}, (_, i) => ['c' + i, 1]), ['root', count]];
		}
		const raw = new Map(entries);
		measure(count + '-' + mode, () => api.resolveRawIntegerConstants(raw), result => assert.deepEqual([...result], expected),
			count === 1 ? 100 : count === 100 ? 10 : 1, Boolean(baseline) && count === 10000 && (mode === 'forward' || mode === 'cycle'));
	}
}
for (const [expression, expected] of [['1', 1], ['name', 2], ['name + 1', 3], ['CInt(3.5) + (name * 2)', 8]]) {
	measure(expression, () => api.evaluateIntegerConstantExpression(expression, {get: () => 2}), result => assert.equal(result, expected), 1000);
}
console.log(JSON.stringify({baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9,
	scope: 'Complete expression evaluation or constant-map resolution, including tokenization, parser allocation and lookups. Input and independent complete result assertions outside timer. Lexer caches warmed. Overflow samples are failed baseline work, not comparable completed-work latency. No editor-latency or heap-byte claim.', rows}, null, 2));
