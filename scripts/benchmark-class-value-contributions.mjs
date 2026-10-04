// Run: node scripts/benchmark-class-value-contributions.mjs [--baseline=COMMIT]
import {build} from 'esbuild';
import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';
import {mkdtempSync, writeFileSync, unlinkSync, rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {cpus, tmpdir} from 'node:os';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-class-value-')), file = join(scratch, 'api.cjs');
let api;
try {
	const plugins = baseline ? [{name: 'baseline', setup(builder) {
		builder.onLoad({filter: /[\\/]symbols[\\/]projectIndex\.ts$/}, args => ({contents: execFileSync('git', ['show', baseline + ':src/analyzer/symbols/projectIndex.ts'], {cwd: root, encoding: 'utf8'}), loader: 'ts', resolveDir: dirname(args.path)}));
	}}] : [];
	const result = await build({plugins, stdin: {contents: "export {ProjectIndex} from './src/analyzer/symbols/projectIndex';", resolveDir: root, loader: 'ts'}, bundle: true, platform: 'node', format: 'cjs', write: false});
	writeFileSync(file, result.outputFiles[0].contents);
	api = createRequire(import.meta.url)(file);
} finally {
	try { unlinkSync(file); } catch (error) { if (error.code !== 'ENOENT') { throw error; } }
	rmdirSync(scratch);
}
const source = 'Public Ref As Object\nPublic EmptyValue As Variant\n' + Array.from({length: 20}, (_, i) => `Public Function Value${i}() As Variant\nDim n As Long\nn = 1\nValue${i} = 42\nEnd Function\n`).join('');
function setup(count) {
	const index = new api.ProjectIndex();
	for (let i = 0; i < count; i++) { index.setModule({moduleName: 'Class' + i, moduleKind: 'class', source}); }
	index.setModule({moduleName: 'Caller', moduleKind: 'standard', source: 'Sub S()\nEnd Sub'});
	return index;
}
const rows = [];
for (const count of [1, 100, 1000]) {
	for (const mode of ['cold', 'after-caller-edit', 'editor-after-caller-edit']) {
		const samples = [];
		let expected;
		for (let round = -3; round < 9; round++) {
			const index = setup(count), options = {includeClassValueFacts: mode !== 'editor-after-caller-edit'};
			if (!expected) { expected = structuredClone(index.projectClassMembers(options)); }
			if (mode !== 'cold') { index.projectClassMembers(options); index.setModule({moduleName: 'Caller', moduleKind: 'standard', source: 'Sub S()\nDim n As Long\nEnd Sub'}); }
			// Cold uses a fresh project that has not queried class surfaces.
			const measured = mode === 'cold' ? setup(count) : index;
			const start = performance.now(), actual = measured.projectClassMembers(options), elapsed = performance.now() - start;
			assert.deepEqual(actual, expected);
			for (const type of actual) { assert.deepEqual(type.members.map(member => [member.name, member.knownValue]), [['Ref', options.includeClassValueFacts ? 'nothing' : undefined], ['EmptyValue', options.includeClassValueFacts ? 'empty' : undefined], ...Array.from({length: 20}, (_, i) => ['Value' + i, options.includeClassValueFacts ? 'scalar' : undefined])]); }
			if (round >= 0) { samples.push(elapsed); }
		}
		samples.sort((a, b) => a - b);
		rows.push({count, mode, classCharacters: source.length, medianMs: samples[4], maxMs: samples[8]});
	}
}
console.log(JSON.stringify({baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, scope: 'Complete projectClassMembers query, including surface materialization, value facts and returned array copy. Parsing/project setup/edit excluded, lexer warmed. Cold means fresh index contributions, not cold lexer/process. Independent value assertions and complete snapshot checks outside timer. No heap-byte or editor-latency claim.', rows}, null, 2));
