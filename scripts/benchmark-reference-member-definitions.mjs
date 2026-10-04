import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, relative } from 'node:path';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';

const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const root = process.cwd(), scratch = mkdtempSync(join(tmpdir(), 'xlide-member-definitions-'));
const path = join(scratch, 'api.cjs');
try {
    const plugins = baseline ? [{ name: 'baseline', setup(builder) {
        builder.onLoad({ filter: /(?:vbaReferenceResolution|memberAccess)\.ts$/ }, args => {
            const file = relative(root, args.path).replaceAll('\\', '/');
            return { contents: execFileSync('git', ['show', baseline + ':' + file], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) };
        });
    } }] : [];
    const bundle = await build({ plugins, stdin: { contents: "export {collectSymbolReferences} from './src/vbaReferenceResolution'; export {buildVbaProjectIndex} from './src/vbaProjectAnalysis'; export {resolveMemberDefinitionsAt} from './src/analyzer/completion/memberAccess';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(path, bundle.outputFiles[0].contents);
    const api = createRequire(import.meta.url)(path), rows = [];
    function measure(configuration, run, validate = () => {}) {
        const times = [];
        let expected;
        for (let i = -3; i < 15; i++) {
            const start = performance.now(), result = run(), elapsed = performance.now() - start;
            validate(result);
            if (i === -3) { expected = result; } else { assert.deepEqual(result, expected); }
            if (i >= 0) { times.push(elapsed); }
        }
        times.sort((a, b) => a - b);
        rows.push({ ...configuration, medianMs: times[7], p95Ms: times[14], resultSha256: createHash('sha256').update(JSON.stringify(expected)).digest('hex') });
    }
    for (const members of [1, 100, 1000]) {
        const library = Array.from({ length: members }, (_, i) => 'Public Sub P' + i + '()\nEnd Sub\n').join('');
        for (const references of [0, 1, 1000]) {
            for (const position of references === 0 ? ['last'] : ['first', 'last']) {
                const name = position === 'first' ? 'P0' : 'P' + (members - 1);
                const modules = [{ moduleName: 'Library', source: library }, { moduleName: 'Caller', source: 'Sub UseIt()\n' + Array(references).fill('    Library.' + name).join('\n') + '\nEnd Sub\n' }];
                const project = api.buildVbaProjectIndex(modules), byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
                const at = library.indexOf(name + '()');
                measure({ scope: 'collector', members, references, position }, () => {
                    const result = api.collectSymbolReferences(byModule, project, modules, library, 'Library', modules[0], name, at + name.length, at, true);
                    return result;
                }, result => {
                    assert.equal(result.hasSymbol, true); assert.equal(result.references.length, references + 1); assert.equal(result.ambiguous.length, 0);
                });
            }
        }
        const project = api.buildVbaProjectIndex([{ moduleName: 'Library', source: library }]);
        const projectClassMembers = project.projectMemberSurfaces('Caller');
        for (const position of ['first', 'last', 'missing']) {
            const name = position === 'first' ? 'P0' : position === 'last' ? 'P' + (members - 1) : 'Missing';
            const source = 'Sub UseIt()\nLibrary.' + name + '\nEnd Sub\n', offset = source.indexOf('Library.' + name) + ('Library.' + name).length;
            const expected = api.resolveMemberDefinitionsAt(source, offset, name, { projectClassMembers });
            assert.equal(expected.length, position === 'missing' ? 0 : 1);
            for (const scope of ['direct-uncached-one', 'direct-cached-one', 'direct-cached-many']) {
                const queries = scope === 'direct-cached-many' ? 1000 : 1;
                measure({ scope, members, references: queries, position }, () => {
                    const ctx = scope === 'direct-uncached-one' ? { projectClassMembers } : { projectClassMembers, memberSurfaceCache: new Map() };
                    let result;
                    for (let i = 0; i < queries; i++) { result = api.resolveMemberDefinitionsAt(source, offset, name, ctx); }
                    return result;
                });
            }
        }
    }
    console.log(JSON.stringify({ baseline: baseline ?? 'working-tree', node: process.version, rounds: 15, warmups: 3, scope: 'Actual complete reference collection or direct definition API; warm project/source caches; new cache context for each direct measurement; complete output comparisons and invariant assertions outside timing; repeated identical direct queries return the final result; no Office/UI timings', rows }, null, 2));
} finally {
    if (existsSync(path)) { unlinkSync(path); }
    rmdirSync(scratch);
}
