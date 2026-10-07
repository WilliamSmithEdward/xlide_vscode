import { describe, expect, it } from 'vitest';
import { moveToModule, type MoveToModuleInput } from '../src/analyzer/refactor/moveToModule';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const eol of ['\n','\r\n','\r'])describe('Move module eligibility '+JSON.stringify(eol),()=>{
 const source=['Public Sub Build()','Debug.Print "built"','End Sub',''].join(eol);
 for(const kind of ['class','document','userform'] as const)for(const role of ['source','target'] as const)it('refuses '+kind+' '+role+' without returning edits',()=>{
  const input:MoveToModuleInput & {moduleKinds:Record<string,typeof kind|'standard'>}={source,offset:source.indexOf('Build'),moduleName:'Reports',targetModuleName:'Helpers',otherModuleSources:{Helpers:''},moduleKinds:{rEpOrTs:role==='source'?kind:'standard',hElPeRs:role==='target'?kind:'standard'}};
  const result=moveToModule(input);
  expect(result.ok).toBe(false);
  if(result.ok)throw Error('unsafe move accepted');
  expect(result.reason).toContain(role==='source'?'Reports':'Helpers');
  expect(result.reason).toMatch(/standard module/i);
  expect(result).not.toHaveProperty('edits');expect(result).not.toHaveProperty('otherModules');
 });
 for(const metadata of [undefined,{rEpOrTs:'standard',hElPeRs:'standard',Caller:'class'}] as const)it('preserves full standard-module output with metadata='+!!metadata,()=>{
  const caller=['Public Sub Caller()','Reports.Build','End Sub',''].join(eol);
  const input:MoveToModuleInput & {moduleKinds?:Record<string,'standard'|'class'>}={source,offset:source.indexOf('Build'),moduleName:'Reports',targetModuleName:'Helpers',otherModuleSources:{Helpers:'',Caller:caller},moduleKinds:metadata};
  const result=moveToModule(input);if(!result.ok)throw Error(result.reason);
  expect(applyVbaTextEdits(source,result.edits)).toBe('');
  expect(applyVbaTextEdits('',result.otherModules!.find(m=>m.moduleName==='Helpers')!.edits)).toBe(eol+source);
  expect(applyVbaTextEdits(caller,result.otherModules!.find(m=>m.moduleName==='Caller')!.edits)).toBe(['Public Sub Caller()','Helpers.Build','End Sub',''].join(eol));
 });
});
