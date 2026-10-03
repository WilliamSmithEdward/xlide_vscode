// Run in the target checkout: node scripts/benchmark-directive-forms.mjs [--rounds=15]
import { buildSync } from 'esbuild';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-directive-forms-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { checkDirectiveForms } from './src/analyzer/diagnostics/rules/directiveForms';
        export { parseModule } from './src/analyzer/parser/parseModule';
        export { indexConditionalCompilation, createConditionalActivityTracker } from './src/analyzer/conditional/conditionalCompilation';
        
        `, resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
const rows = [];
function measure(name, run) {
    for (let i = 0; i < 3; i++) run();
    const samples = [];
    for (let i = 0; i < rounds; i++) {
        const start = performance.now(); run(); samples.push(performance.now() - start);
    }
    samples.sort((a,b)=>a-b);
    rows.push({ name, medianMs: +samples[Math.floor(samples.length/2)].toFixed(3),
        p95Ms: +samples[Math.ceil(samples.length*.95)-1].toFixed(3) });
}
for(const count of [100,1000,5000]) {
    const source=Array.from({length:count},(_,i)=>'#Const FEATURE'+i+' = 1\n').join('');
    const module=api.parseModule(source);
    measure(count+'-const-directives',()=>{let found=0;api.checkDirectiveForms(source,module,undefined,()=>{found++;});if(found!==0)throw new Error('Unexpected diagnostic');});
}
const source=Array.from({length:1000},(_,i)=>'#Const FEATURE'+i+' = 1: Debug.Print 1\n').join('');
const module=api.parseModule(source);
measure('1000-trailing-statements',()=>{let found=0;api.checkDirectiveForms(source,module,undefined,kind=>{if(kind==='directiveTrailingStatement')found++;});if(found!==1000)throw new Error('Missing diagnostics: '+found);});
const constantsSource=Array.from({length:1000},(_,i)=>'#Const FLAG'+i+' = '+(i===0?'1':'FLAG'+(i-1)+' + 1')+'\n').join('')+'#If FLAG999 = 1000 Then\nDebug.Print 1\n#End If\n';
const constantsModule=api.parseModule(constantsSource);
measure('1000-dependent-constant-index',()=>{if(api.indexConditionalCompilation(constantsModule).constants[999].value!==1000)throw new Error('Wrong constant value');});
measure('1000-dependent-constant-tracker',()=>{const tracker=api.createConditionalActivityTracker(constantsModule);const offset=constantsSource.indexOf('Debug.Print');if(tracker.activityForSpan({start:offset,end:offset+1})!=='active')throw new Error('Wrong branch activity');});
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
