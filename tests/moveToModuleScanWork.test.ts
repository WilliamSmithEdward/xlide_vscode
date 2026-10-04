import { expect, it, vi } from 'vitest';
const work=vi.hoisted(()=>({scannedCharacters:0}));
vi.mock('../src/vbaSourceScan',async original=>{const actual=await original<typeof import('../src/vbaSourceScan')>();return {...actual,stripVba:(source:string)=>{work.scannedCharacters+=source.length;return actual.stripVba(source);}};});
import { moveToModule } from '../src/analyzer/refactor/moveToModule';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const count of [1,100,1000]) for(const multiline of [false,true]) it('scans each matched physical line once at '+count+', multiline='+multiline,()=>{
 const source='Public Sub Build()\nEnd Sub\n';
 const calls=Array(count).fill('Reports.Build');
 const caller=['Sub Caller()',calls.join(multiline?'\n':': '),'End Sub',''].join('\n');
 work.scannedCharacters=0;
 const result=moveToModule({source,offset:source.indexOf('Build'),moduleName:'Reports',targetModuleName:'Helpers',otherModuleSources:{Helpers:'',Caller:caller}});
 if(!result.ok)throw new Error(result.reason);
 expect(work.scannedCharacters).toBeLessThanOrEqual(source.length+caller.length);
 const changes=result.otherModules!.find(module=>module.moduleName==='Caller')!;
 expect(applyVbaTextEdits(caller,changes.edits)).toBe(['Sub Caller()',Array(count).fill('Helpers.Build').join(multiline?'\n':': '),'End Sub',''].join('\n'));
});
