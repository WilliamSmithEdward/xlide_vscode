import {afterEach,describe,expect,it,vi} from 'vitest';
import * as parser from '../src/analyzer/parser/parseModule';
import * as host from '../src/analyzer/host/hostModel';
import {resolveEventHandlerCompletions,type EventHandlerCompletionContext} from '../src/analyzer/completion/eventHandlers';
import type {HostObjectModel} from '../src/analyzer/host/excelObjectModel';
afterEach(()=>vi.restoreAllMocks());
const model:HostObjectModel={source:'Synthetic event work fixture',types:{'VB.Form':{displayName:'Form',members:[{name:'Load',kind:'event',signature:'Load()',doc:{summary:'Form loads.'}}]},'Access.Form':{displayName:'Form',members:[{name:'Load',kind:'event',signature:'Load()',doc:{summary:'Form loads.'}}]},'VB.Button':{displayName:'Button',members:[{name:'Click',kind:'event',signature:'Click()',doc:{summary:'Control clicks.'}}]}},aliases:{},globals:{}};
const controls=Array.from({length:1000},(_,i)=>({name:'Button'+i,type:'VB.Button'}));
const contexts:EventHandlerCompletionContext[]=[{moduleKind:'document',documentType:'workbook'},{moduleKind:'userform',host:'vb6',meType:'VB.Form',model,implicitMembers:controls},{moduleKind:'userform',meType:'Access.Form',model,implicitMembers:controls}];
describe('event completion prefix gates',()=>{
 for(const eol of ['\n','\r\n','\r'])it.each(contexts)('rejects nonmatching prefixes before parsing (%j, '+JSON.stringify(eol)+')',ctx=>{
  const source='Sub Existing()'+eol+'End Sub'+eol+'UnrelatedPrefix';
  const parse=vi.spyOn(parser,'parseModule'),events=vi.spyOn(host,'getHostEvents');
  expect(resolveEventHandlerCompletions(source,source.length,ctx)).toEqual([]);
  expect(parse).not.toHaveBeenCalled();expect(events).not.toHaveBeenCalled();
 });
 it.each(contexts.slice(1))('looks up only the matching control owner (%j)',ctx=>{
  const source='Button999_C',events=vi.spyOn(host,'getHostEvents');
  expect(resolveEventHandlerCompletions(source,source.length,ctx)).toEqual([{name:'Button999_Click',signature:'Button999_Click()',detail:'Button event handler',documentation:'Control clicks.',insertText:'Private Sub Button999_Click()\n    $0\nEnd Sub'}]);
  expect(events).toHaveBeenCalledTimes(1);
 });
 it('retains existing-handler exclusion and reads changing designer context',()=>{
  const ctx=contexts[1],source='Sub Button999_Click()\nEnd Sub\nButton999_C';
  expect(resolveEventHandlerCompletions(source,source.length,ctx)).toEqual([]);
  const prefix='NewControl_C',list=[{name:'Button',type:'VB.Button'}],changed={...ctx,implicitMembers:list};
  expect(resolveEventHandlerCompletions(prefix,prefix.length,changed)).toEqual([]);
  list.push({name:'NewControl',type:'VB.Button'});
  expect(resolveEventHandlerCompletions(prefix,prefix.length,changed).map(r=>r.name)).toEqual(['NewControl_Click']);
 });
});
