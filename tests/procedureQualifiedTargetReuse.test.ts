import { expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import { parseModule } from '../src/analyzer/parser/parseModule';
import type { ProcedureNode } from '../src/analyzer/parser/nodes';
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { procedureCallBinding } from '../src/analyzer/refactor/procedureCallBinding';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

for (const count of [1, 1000]) for (const access of ['Public', 'Private']) {
 it('resolves the owning declaration once for '+count+' qualified '+access+' calls', () => {
  const source=Array.from({length:100},(_,i)=>'Sub Earlier'+i+'()\nEnd Sub\n').join('')+access+' Sub Report()\nDim limit As Long\nlimit = 3\nDebug.Print limit\nEnd Sub\n';
  const caller='Sub Caller()\n'+Array(count).fill('Module1.Report').join('\n')+'\nEnd Sub\n';
  const project=new ProjectIndex();
  project.setModule({moduleName:'Module1',moduleKind:'standard',source});
  project.setModule({moduleName:'Caller',moduleKind:'standard',source:caller});
  const procedure=parseModule(source).members.find(m=>m.kind==='Procedure'&&m.name==='Report') as ProcedureNode;
  const search=vi.spyOn(project.getModule('Module1')!.root.children!, 'find');
  try {
   const accepts=procedureCallBinding(source,'Module1',procedure,{Caller:caller},()=>project);
   expect(callSitesOf(caller,'Report').map(site=>accepts('Caller',caller,site))).toEqual(Array(count).fill(access==='Public'));
   expect(search).toHaveBeenCalledTimes(1);
  } finally { search.mockRestore(); }
 });
}
it('caches a missing owning declaration too',()=>{
 const source='Public Sub Report()\nEnd Sub\n', caller='Sub Caller()\nModule1.Report\nModule1.Report\nEnd Sub\n';
 const project=new ProjectIndex();project.setModule({moduleName:'Module1',moduleKind:'standard',source});project.setModule({moduleName:'Caller',moduleKind:'standard',source:caller});
 const procedure=parseModule(source).members[0] as ProcedureNode;
 const search=vi.spyOn(project.getModule('Module1')!.root.children!,'find');
 try {const accepts=procedureCallBinding(source,'Module1',{...procedure,span:{...procedure.span,start:9999}},{Caller:caller},()=>project);expect(callSitesOf(caller,'Report').map(s=>accepts('Caller',caller,s))).toEqual([false,false]);expect(search).toHaveBeenCalledTimes(1);}finally{search.mockRestore();}
});
for(const eol of ['\n','\r\n','\r'])it('keeps complete edits, shadowed receivers and fresh query visibility '+JSON.stringify(eol),()=>{
 const base=['Sub Earlier()','End Sub','Public Sub Report()','Dim limit As Long','limit = 3','Debug.Print limit','End Sub','Sub LocalCaller()','Module1.Report','End Sub',''].join(eol);
 const caller=['Sub Caller()','Module1.Report','Call [mOdUlE1].[Report]()','End Sub','Sub Shadowed(ByVal Module1 As Object)','Module1.Report','End Sub',''].join(eol);
 for(const access of ['Public','Private','Public']){
  const source=base.replace('Public Sub Report',access+' Sub Report');const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',otherModuleSources:{Caller:caller}});if(!result.ok)throw Error(result.reason);
  expect(applyVbaTextEdits(source,result.edits)).toBe(['Sub Earlier()','End Sub',access+' Sub Report(ByVal limit As Long)','Debug.Print limit','End Sub','Sub LocalCaller()','Module1.Report 3','End Sub',''].join(eol));
  expect(result.otherModules??[]).toEqual(access==='Private'?[]:[{moduleName:'Caller',edits:[{span:{start:caller.indexOf('Module1.Report')+'Module1.Report'.length,end:caller.indexOf('Module1.Report')+'Module1.Report'.length},newText:' 3'},{span:{start:caller.indexOf(']()')+2,end:caller.indexOf(']()')+2},newText:'3'}]}]);
 }
});
