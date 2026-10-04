import { describe, expect, it, vi } from 'vitest';
const state = vi.hoisted(() => ({handlers:new Map<string,()=>unknown>(), edits:[] as Array<{start:number,end:number,text:string}>}));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
 commands:{registerCommand:vi.fn((name:string,handler:()=>unknown)=>{state.handlers.set(name,handler);return {dispose:vi.fn()};})},
 window:{activeTextEditor:undefined,setStatusBarMessage:vi.fn()},
 workspace:{applyEdit:vi.fn(async()=>true)},
 WorkspaceEdit:class {replace(_uri:unknown,range:{start:{character:number},end:{character:number}},text:string){state.edits.push({start:range.start.character,end:range.end.character,text});}},
}));
vi.mock('../src/projectModuleOperations',()=>({writeProjectModule:vi.fn(async()=>({ok:true})),refreshProjectState:vi.fn()}));
import * as vscode from 'vscode';
import { registerRefactorCommands } from '../src/commands/refactorCommands';
import { writeProjectModule, refreshProjectState } from '../src/projectModuleOperations';
import { encodeModuleUri } from '../src/xlideFileSystem';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

for(const eol of ['\n','\r\n','\r']){
 describe('Introduce Parameter command binding '+JSON.stringify(eol),()=>{
  it('writes the owning call without editing the unrelated module-qualified call',async()=>{
   vi.clearAllMocks();state.handlers.clear();state.edits.length=0;
   const path=process.platform==='win32'?'C:\\work\\Report.xlsm':'/work/Report.xlsm';
   const source=['Public Sub Report()','Dim limit As Long','limit = 3','Debug.Print limit','End Sub',''].join(eol);
   const caller=['Sub Caller()','Module1.Report','Module2.Report','End Sub',''].join(eol);
   const other=['Public Sub Report()','Debug.Print "other"','End Sub',''].join(eol);
   const sources:Record<string,string>={Caller:caller,Module2:other};
   const bridge={call:vi.fn(async(method:string,params:{module?:string})=>{
    if(method==='listModules')return [{name:'Module1'},{name:'Caller'},{name:'Module2'}];
    if(method==='readModule'&&params.module&&sources[params.module])return {source:sources[params.module]};
    throw new Error('unexpected bridge call '+method);
   })};
   const deps={bridge,explorer:{},fsProvider:{},vbaIndex:{},out:{},context:{}};
   const at={line:0,character:source.indexOf('limit')};
   (vscode.window as {activeTextEditor:unknown}).activeTextEditor={document:{uri:encodeModuleUri(path,'Module1'),languageId:'vba',getText:()=>source,offsetAt:()=>at.character,positionAt:(offset:number)=>({line:0,character:offset})},selection:{active:at,start:at,end:at},revealRange:vi.fn()};
   registerRefactorCommands(deps as never);
   await state.handlers.get('xlide.refactor.introduceParameter')!();
   expect(applyVbaTextEdits(source,state.edits.map(edit=>({span:{start:edit.start,end:edit.end},newText:edit.text})))).toBe(['Public Sub Report(ByVal limit As Long)','Debug.Print limit','End Sub',''].join(eol));
   expect(writeProjectModule).toHaveBeenCalledTimes(1);
   expect(writeProjectModule).toHaveBeenCalledWith(deps,{filePath:path,moduleName:'Caller',source:['Sub Caller()','Module1.Report 3','Module2.Report','End Sub',''].join(eol)},{refreshProjectState:false});
   expect(refreshProjectState).toHaveBeenCalledTimes(1);
   expect(bridge.call).not.toHaveBeenCalledWith('writeModule',expect.anything());
  });
 });
}
