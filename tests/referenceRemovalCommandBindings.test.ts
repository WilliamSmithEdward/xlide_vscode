import {expect,it,vi,beforeEach} from 'vitest';
const handlers=vi.hoisted(()=>new Map<string,(...args:any[])=>Promise<unknown>>());
vi.mock('vscode',async()=> (await import('./helpers/vscodeMock')).vscodeMock());
vi.mock('../src/xlideCommandRegistration',()=>({registerXlideCommand:(name:string,handler:any)=>{handlers.set(name,handler);return {dispose(){}};}}));
vi.mock('../src/officeWriteCoordinator',()=>({runWriteWithHostCoordination:async(_path:string,action:()=>unknown)=>action()}));
vi.mock('../src/projectModuleOperations',()=>({refreshProjectState:vi.fn(),deleteProjectModule:vi.fn(),renameProjectModule:vi.fn(),writeProjectModule:vi.fn()}));
vi.mock('../src/xlideWriteAudit',()=>({recordXlideWriteAuditEvent:vi.fn()}));
import * as vscode from 'vscode';
import {registerProjectCrudCommands} from '../src/commands/projectCrudCommands';
import type {CommandDeps} from '../src/commands/shared';
beforeEach(()=>{vi.clearAllMocks();handlers.clear();});
function setup(modules:any[]){
 const call=vi.fn(async()=>({removed:true,name:'Word'}));
 registerProjectCrudCommands({bridge:{call},vbaIndex:{getAllModules:vi.fn(async()=>modules)},out:{appendLine:vi.fn()},explorer:{refresh:vi.fn()},fsProvider:{}} as unknown as CommandDeps);
 return call;
}
it('does not warn for exported source values named as a library',async()=>{
 const call=setup([{moduleName:'Globals',source:'Public Word As Worksheet'},{moduleName:'Caller',source:'Sub T()\nWord.EnableCalculation = True\nEnd Sub'}]);
 await handlers.get('xlide.removeProjectReference')!('/test.xlsm','Word');
 expect(vscode.window.showWarningMessage).not.toHaveBeenCalled();
 expect(call).toHaveBeenCalledWith('removeReference',{path:'/test.xlsm',library:'Word'});
});
it('requires the existing warning for a genuine early-bound dependency',async()=>{
 const call=setup([{moduleName:'Caller',source:'Dim app As Word.Application'}]);
 await handlers.get('xlide.removeProjectReference')!('/test.xlsm','Word');
 expect(vscode.window.showWarningMessage).toHaveBeenCalled();
 expect(call).not.toHaveBeenCalled();
});
