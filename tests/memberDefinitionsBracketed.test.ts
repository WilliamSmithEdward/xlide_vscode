import { describe, expect, it } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { resolveMemberDefinitionsAt, type MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
for(const eol of ['\n','\r\n','\r']) {
 describe('bracketed member definition '+JSON.stringify(eol),()=>{
  for(const body of ['Me.[Report] 1','Dim value As Target'+eol+'value.[Report] 1','With Me'+eol+'.[Report] 1'+eol+'End With']) {
   it(body,()=>{
    const target=['Public Sub Report(ByVal value As Long)','End Sub',''].join(eol);
    const source=['Sub Caller()',body,'End Sub',''].join(eol);
    const project=new ProjectIndex();project.setModule({moduleName:'Target',moduleKind:'class',source:target});project.setModule({moduleName:'Caller',moduleKind:'class',source});
    const context:MemberCompletionContext={meProjectType:'Target',projectClassMembers:project.projectMemberSurfaces('Caller'),sourceTokens:tokenizeCached(source).filter(t=>t.kind!=='comment'),withScanCache:new Map(),receiverTypeCache:new Map(),receiverChainCache:new Map(),memberSurfaceCache:new Map()};
    const actual=resolveMemberDefinitionsAt(source,source.indexOf('[Report]')+'[Report]'.length,'Report',context);
    expect(actual).toEqual([{moduleName:'Target',nameSpan:{start:target.indexOf('Report'),end:target.indexOf('Report')+6},fullSpan:{start:0,end:target.trimEnd().length}}]);
    const plain=source.replace('[Report]','Report');
    expect(resolveMemberDefinitionsAt(plain,plain.lastIndexOf('Report')+6,'Report',{...context,sourceTokens:tokenizeCached(plain).filter(t=>t.kind!=='comment'),withScanCache:new Map(),receiverTypeCache:new Map(),receiverChainCache:new Map(),memberSurfaceCache:new Map()})).toEqual(actual);
   });
  }
 });
}
