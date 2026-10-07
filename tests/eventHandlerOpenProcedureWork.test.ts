import {describe,expect,it} from 'vitest';
import {parseModule} from '../src/analyzer/parser/parseModule';
import {resolveEventHandlerCompletions} from '../src/analyzer/completion/eventHandlers';
const ctx={moduleKind:'document' as const,documentType:'workbook' as const};
const expected=[{name:'Workbook_Open',signature:'Workbook_Open()',detail:'Workbook event handler',documentation:'Occurs when the workbook is opened.',insertText:'Private Sub Workbook_Open()\n    $0\nEnd Sub'}];
describe('open procedure event completion work',()=>{
 for(const eol of ['\n','\r\n','\r'])it.each([10,100,1000])('bounds repeated open-procedure boundary reads (%i, '+JSON.stringify(eol)+')',count=>{
  const source=Array.from({length:count},(_,i)=>'Sub Broken'+i+'()'+eol).join('')+'Sub Closed()'+eol+'End Sub'+eol+'Workbook_Open';
  const mod=parseModule(source);expect(mod.members.filter(m=>m.kind==='Procedure'&&!m.closed)).toHaveLength(count);
  let reads=0;const restore:Array<()=>void>=[];
  for(const member of mod.members){const span=member.span,descriptor=Object.getOwnPropertyDescriptor(span,'start')!,value=span.start;Object.defineProperty(span,'start',{get(){reads++;return value;},configurable:true});restore.push(()=>Object.defineProperty(span,'start',descriptor));}
  try{expect(resolveEventHandlerCompletions(source,source.length,ctx)).toEqual(expected);expect(reads).toBeLessThanOrEqual(6*mod.members.length);}finally{for(const reset of restore)reset();}
 });
 it.each(['\n','\r\n','\r'])('retains refusal inside the first/last open procedure and allows the following closed boundary (%j)',eol=>{
  const source=['Sub First()','Workbook_Open','Sub Last()','Workbook_Open','Sub Closed()','End Sub','Workbook_Open'].join(eol);
  for(const offset of [source.indexOf('Workbook_Open')+13,source.indexOf('Workbook_Open',source.indexOf('Workbook_Open')+13)+13])expect(resolveEventHandlerCompletions(source,offset,ctx)).toEqual([]);
  expect(resolveEventHandlerCompletions(source,source.length,ctx)).toEqual(expected);
 });
});
