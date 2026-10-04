import { describe, expect, it, vi } from 'vitest';
const state=vi.hoisted(()=>({handlers:new Map<string,()=>unknown>()}));
vi.mock('vscode',async()=>(await import('./helpers/vscodeMock')).vscodeMock({
 commands:{registerCommand:vi.fn((name:string,handler:()=>unknown)=>{state.handlers.set(name,handler);return{dispose:vi.fn()};})},
 window:{activeTextEditor:undefined,showQuickPick:vi.fn(async()=> 'Helpers'),setStatusBarMessage:vi.fn()},
 workspace:{applyEdit:vi.fn(async()=>true)},
 WorkspaceEdit:class {replace():void{}},
}));
vi.mock('../src/projectModuleOperations',()=>({writeProjectModule:vi.fn(async()=>({ok:true})),refreshProjectState:vi.fn()}));
import * as vscode from 'vscode';
import { registerRefactorCommands } from '../src/commands/refactorCommands';
import { writeProjectModule, refreshProjectState } from '../src/projectModuleOperations';
import { encodeModuleUri } from '../src/xlideFileSystem';
async function run(sourceType?:string,targetType?:string,extra:Array<{name:string,type:string}>=[]){
 vi.clearAllMocks();state.handlers.clear();
 const path=process.platform==='win32'?'C:\\work\\Reports.xlsm':'/work/Reports.xlsm';
 const source='Public Sub Build()\nEnd Sub\n';const at={line:0,character:source.indexOf('Build')};
 (vscode.window as {activeTextEditor:unknown}).activeTextEditor={document:{uri:encodeModuleUri(path,'Reports'),languageId:'vba',getText:()=>source,offsetAt:()=>at.character,positionAt:(character:number)=>({line:0,character})},selection:{active:at,start:at,end:at},revealRange:vi.fn()};
 const bridge={call:vi.fn(async(method:string)=>{if(method==='listModules')return[{name:'Reports',type:sourceType},{name:'Helpers',type:targetType},...extra];if(method==='readModule')return{source:''};throw Error(method);})};
 registerRefactorCommands({bridge,explorer:{},fsProvider:{},vbaIndex:{},out:{},context:{}} as never);
 await state.handlers.get('xlide.refactor.moveToModule')!();
}
describe('Move command module roles',()=>{
 for(const type of ['class','document','userform','usercontrol','propertypage','designer','accessform','accessreport'])for(const role of ['source','target'])it('does not write '+type+' '+role,async()=>{
  await run(role==='source'?type:'standard',role==='target'?type:'standard');
  expect(vscode.workspace.applyEdit).not.toHaveBeenCalled();expect(writeProjectModule).not.toHaveBeenCalled();expect(refreshProjectState).not.toHaveBeenCalled();
  expect(vscode.window.showInformationMessage).toHaveBeenCalledWith(expect.stringMatching(/standard module/i));
 });
 it('offers only standard targets and writes a standard destination',async()=>{
  await run('standard','standard',[{name:'Class1',type:'class'},{name:'Sheet1',type:'document'},{name:'Form1',type:'userform'}]);
  expect(vscode.window.showQuickPick).toHaveBeenCalledWith(['Helpers'],expect.anything());
  expect(vscode.workspace.applyEdit).toHaveBeenCalledTimes(1);expect(writeProjectModule).toHaveBeenCalledTimes(1);
 });
 it('preserves the older unspecified module-type default',async()=>{await run();expect(vscode.workspace.applyEdit).toHaveBeenCalledTimes(1);expect(writeProjectModule).toHaveBeenCalledTimes(1);});
});
