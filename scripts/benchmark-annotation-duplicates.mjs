import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const root = process.cwd(), scratch = mkdtempSync(join(tmpdir(), 'xlide-annotation-duplicates-'));
const path = join(scratch, 'api.cjs');
try {
    const plugins = baseline ? [{ name: 'baseline', setup(builder) {
        builder.onLoad({ filter: /attributeAnnotations\.ts$/ }, args => ({
            contents: execFileSync('git', ['show', baseline + ':src/analyzer/annotations/attributeAnnotations.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path),
        }));
    } }] : [];
    const bundle = await build({ plugins, stdin: { contents: "export {readAttributeAnnotations} from './src/analyzer/annotations/attributeAnnotations'; export {applyAttributeAnnotations} from './src/analyzer/annotations/attributeRewriter';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(path, bundle.outputFiles[0].contents);
    const api = createRequire(import.meta.url)(path), rows = [];
    for (const count of [1, 100, 1000]) {
        const declarations = Array.from({ length: count }, (_, i) => 'Sub P' + i + '()\nEnd Sub\n').join('');
        const variableDeclarations = Array.from({ length: count }, (_, i) => "'@VariableDescription(\"desc\")\n" + 'Public v' + i + ' As Long\n').join('');
        const sources = {
            procedures: Array.from({ length: count }, (_, i) => "'@Description(\"desc\")\n" + 'Sub P' + i + '()\nEnd Sub\n').join(''),
            variables: variableDeclarations,
            'procedure-duplicates': Array(count).fill("'@Description(\"desc\")\n").join('') + 'Sub P()\nEnd Sub\n',
            'module-duplicates': variableDeclarations + Array(count).fill("'@ModuleDescription(\"module\")\n").join(''),
            'module-only': "'@ModuleDescription(\"module\")\n" + declarations,
            'unknown-only': "'@Folder(\"unknown\")\n" + declarations,
            'no-annotations': declarations,
        };
        for (const shape of Object.keys(sources)) {
            const source = 'Attribute VB_Name = "M"\n' + sources[shape];
            for (const scope of ['reader', 'reader-writer']) {
                const times = [];
                let expected;
                for (let i = -3; i < 15; i++) {
                    const start = performance.now(), read = api.readAttributeAnnotations(source);
                    const result = scope === 'reader' ? read : { read, rewrite: read.annotations.length > 0 ? api.applyAttributeAnnotations(source, read) : undefined };
                    const elapsed = performance.now() - start;
                    const expectedAnnotations = shape === 'module-duplicates' ? count + 1
                        : shape === 'procedures' || shape === 'variables' ? count
                        : shape === 'unknown-only' || shape === 'no-annotations' ? 0 : 1;
                    const expectedProblems = shape === 'procedure-duplicates' || shape === 'module-duplicates' ? count - 1 : 0;
                    assert.equal(read.annotations.length, expectedAnnotations);
                    assert.equal(read.problems.length, expectedProblems);
                    if (scope === 'reader-writer' && result.rewrite) {
                        assert.equal(result.rewrite.changes.length, expectedAnnotations);
                        assert.equal(result.rewrite.skipped.length, 0);
                    }
                    if (i === -3) { expected = result; } else { assert.deepEqual(result, expected); }
                    if (i >= 0) { times.push(elapsed); }
                }
                times.sort((a, b) => a - b);
                rows.push({ count, shape, scope, medianMs: times[7], p95Ms: times[14], resultSha256: createHash('sha256').update(JSON.stringify(expected)).digest('hex') });
            }
        }
    }
    console.log(JSON.stringify({ baseline: baseline ?? 'working-tree', node: process.version, rounds: 15, warmups: 3, scope: 'Actual annotation reader and optionally guarded attribute writer APIs; LF sources prepared outside timer; complete output comparisons outside timer; excludes save IO/message formatting and Office/UI execution', rows }, null, 2));
} finally {
    if (existsSync(path)) { unlinkSync(path); }
    rmdirSync(scratch);
}
