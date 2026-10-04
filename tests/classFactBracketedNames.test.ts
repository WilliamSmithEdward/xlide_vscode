import {expect,it} from 'vitest';
import {classMemberValues} from '../src/analyzer/symbols/classMemberFacts';
import {buildModuleSymbols} from '../src/analyzer/symbols/buildModuleSymbols';
import {ProjectIndex} from '../src/analyzer/symbols/projectIndex';
import {analyzeModule} from '../src/analyzer/diagnostics/analyzeModule';
const facts=(source:string)=>[...classMemberValues(source,buildModuleSymbols('C','class',source).root.children??[])];
for(const name of ['member','Прибор','If']){
 it('does not infer Nothing after a bracketed field assignment: '+name,()=>expect(facts('Public ['+name+'] As Object\nSub Setup()\nSet Me.['+name+'] = New Collection\nEnd Sub\n')).toEqual([]));
 it('does not infer Empty after a bracketed Variant assignment: '+name,()=>expect(facts('Public ['+name+'] As Variant\nSub Setup()\nMe.['+name+'] = 1\nEnd Sub\n')).toEqual([]));
 it('recognizes a bracketed object function result: '+name,()=>expect(facts('Public Function ['+name+']() As Object\nSet ['+name+'] = New Collection\nEnd Function\n')).toEqual([]));
 it('recognizes a bracketed scalar function result: '+name,()=>expect(facts('Public Function ['+name+']() As Variant\n['+name+'] = 1\nEnd Function\n')).toEqual([[name.toLowerCase(),'scalar']]));
 it('recognizes a bracketed object getter result: '+name,()=>expect(facts('Public Property Get ['+name+']() As Object\nSet ['+name+'] = New Collection\nEnd Property\n')).toEqual([]));
 it('recognizes a bracketed scalar getter result: '+name,()=>expect(facts('Public Property Get ['+name+']() As Variant\n['+name+'] = 1\nEnd Property\n')).toEqual([[name.toLowerCase(),'scalar']]));
 it('retains a bracketed Nothing result: '+name,()=>expect(facts('Public Function ['+name+']() As Object\nSet ['+name+'] = Nothing\nEnd Function\n')).toEqual([[name.toLowerCase(),'nothing']]));
}
for(const eol of ['\n','\r\n','\r'])for(const kind of ['field','function','getter'] as const)it('preserves complete diagnostics for bracketed '+kind+' '+JSON.stringify(eol),()=>{
 const source=kind==='field'?['Public member As Object','Private Sub Class_Initialize()','Set Me.[member] = New Collection','End Sub',''].join(eol):[kind==='getter'?'Public Property Get member() As Object':'Public Function member() As Object','Set [member] = New Collection',kind==='getter'?'End Property':'End Function',''].join(eol);
 const caller=['Option Explicit','Sub Go()','Dim c As New C','Debug.Print c.member.Count','End Sub',''].join(eol),project=new ProjectIndex();project.setModule({moduleName:'C',moduleKind:'class',source});project.setModule({moduleName:'M',moduleKind:'standard',source:caller});
 const errors:unknown[]=[];expect(analyzeModule(caller,{projectClassMembers:project.projectClassMembers(),onInternalError:e=>errors.push(e)})).toEqual([]);expect(errors).toEqual([]);
});
it.each([['Public member As Object\n',[['member','nothing']]],['Public member As Variant\n',[['member','empty']]],['Public Function member() As Variant\nmember = 1\nEnd Function\n',[['member','scalar']]],['Public member As Object\nSub Go()\nDebug.Print "[member]" \' [member]\nEnd Sub\n',[['member','nothing']]]])('keeps established fact for %s',(source,expected)=>expect(facts(source)).toEqual(expected));
for(const eol of ['\n','\r\n','\r'])for(const mode of ['read','assigned','escaped'] as const)it('handles bracketed instance/member '+mode+' '+JSON.stringify(eol),()=>{
 const operation=mode==='read'?'Debug.Print [c].[member].Count':mode==='assigned'?'Set [c].[member] = New Collection'+eol+'Debug.Print c.member.Count':'Take [c]'+eol+'Debug.Print c.member.Count';
 const caller=['Option Explicit','Sub Go()','Dim c As New C',operation,'End Sub','Sub Take(ByRef value As C)','Set value.member = New Collection','End Sub',''].join(eol),project=new ProjectIndex();project.setModule({moduleName:'C',moduleKind:'class',source:'Public member As Object'+eol});project.setModule({moduleName:'M',moduleKind:'standard',source:caller});
 const failures:unknown[]=[];const diagnostics=analyzeModule(caller,{projectClassMembers:project.projectClassMembers(),onInternalError:e=>failures.push(e)});expect(failures).toEqual([]);
 if(mode==='read'){
  const start=caller.indexOf('[c].[member]');expect(diagnostics).toEqual([{code:'object-variable-not-set',message:"'[c].[member]' is Nothing here: nothing in C sets member. This will raise Run-time error '91': Object variable or With block variable not set.",severity:'error',span:{start,end:start+'[c].[member]'.length},specReference:'VBE runtime error 91: Object variable or With block variable not set',origin:'run'}]);
 }else expect(diagnostics).toEqual([]);
});
