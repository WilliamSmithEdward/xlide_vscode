// Run: node scripts/benchmark-constant-not-chains.mjs [--baseline=COMMIT]
import { build } from 'esbuild';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { cpus, tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { performance } from 'node:perf_hooks';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-constant-not-'));
const file = join(scratch, 'api.cjs');
let api;
try {
	const plugins = baseline ? [{
		name: 'baseline',
		setup(builder) {
			builder.onLoad({ filter: /integerConstantExpression\.ts$/ }, args => ({
				contents: execFileSync('git', ['show', `${baseline}:src/analyzer/constants/integerConstantExpression.ts`], { cwd: root, encoding: 'utf8' }),
				loader: 'ts', resolveDir: dirname(args.path),
			}));
		},
	}] : [];
	const built = await build({
		plugins,
		stdin: { contents: "export {evaluateIntegerConstantExpression} from './src/analyzer/constants/integerConstantExpression';", resolveDir: root, loader: 'ts' },
		bundle: true, platform: 'node', format: 'cjs', write: false,
	});
	writeFileSync(file, built.outputFiles[0].contents);
	api = createRequire(import.meta.url)(file);
} finally {
	try { unlinkSync(file); } catch (error) { if (error.code !== 'ENOENT') throw error; }
	rmdirSync(scratch);
}

const constants = new Map();
const rows = [];
for (const count of [0, 1, 10, 100, 1000, 20000, 20001]) {
	const source = 'Not '.repeat(count) + '7';
	const batchSize = count >= 20000 ? 1 : 100;
	const samples = [];
	let overflows = 0;
	for (let round = -3; round < 9; round++) {
		const start = performance.now();
		let result;
		let overflow = false;
		try {
			for (let i = 0; i < batchSize; i++) result = api.evaluateIntegerConstantExpression(source, constants);
		} catch (error) {
			if (!(error instanceof RangeError) || !/call stack/i.test(error.message)) throw error;
			overflow = true;
		}
		const elapsed = (performance.now() - start) / batchSize;
		if (!overflow) assert.equal(result, count % 2 ? -8 : 7);
		if (round >= 0) {
			if (overflow) overflows++; else samples.push(elapsed);
		}
	}
	samples.sort((a, b) => a - b);
	rows.push({ count, batchSize, overflows, medianMs: samples.length ? samples[Math.floor(samples.length / 2)] : null, maxMs: samples.at(-1) ?? null });
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, scope: 'Whole constant expression evaluator including lexing. Per-call batch averages, output assertions excluded. Overflow samples are failures, not latency results. No whole-analyzer or editor latency claim.', rows }, null, 2));
