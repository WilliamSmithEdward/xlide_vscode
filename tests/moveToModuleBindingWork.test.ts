import { expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import { moveToModule } from '../src/analyzer/refactor/moveToModule';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for (const count of [1,100,1000]) it('shares one project index for '+count+' qualified references per caller', () => {
 const source='Public Sub Build()\nEnd Sub\n';
 const caller='Sub Caller()\n'+Array(count).fill('Reports.Build').join(': ')+'\nEnd Sub\n';
 const spy=vi.spyOn(ProjectIndex.prototype,'setModule');
 try {
  const result=moveToModule({source,offset:source.indexOf('Build'),moduleName:'Reports',targetModuleName:'Helpers',otherModuleSources:{Helpers:'',First:caller,Second:caller}});
  if(!result.ok)throw new Error(result.reason);
  expect(spy.mock.calls.map(([input])=>input.moduleName).sort()).toEqual(['First','Helpers','Reports','Second']);
  for(const name of ['First','Second'])expect(applyVbaTextEdits(caller,result.otherModules!.find(module=>module.moduleName===name)!.edits)).toBe('Sub Caller()\n'+Array(count).fill('Helpers.Build').join(': ')+'\nEnd Sub\n');
 }finally {spy.mockRestore();}
});
it('does not build bindings for comments, strings or unrelated receivers', () => {
 const source='Public Sub Build()\nEnd Sub\n';
 const caller='Sub Caller()\nDebug.Print "Reports.Build"\nRem Reports.Build\nitem.Reports.Build\nEnd Sub\n';
 const spy=vi.spyOn(ProjectIndex.prototype,'setModule');
 try {const result=moveToModule({source,offset:source.indexOf('Build'),moduleName:'Reports',targetModuleName:'Helpers',otherModuleSources:{Helpers:'',Caller:caller}});expect(result.ok).toBe(true);expect(spy).not.toHaveBeenCalled();}finally{spy.mockRestore();}
});
