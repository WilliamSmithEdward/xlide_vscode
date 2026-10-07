import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host=vi.hoisted(()=>({findFiles:vi.fn()}));
vi.mock('vscode',async()=> (await import('./helpers/vscodeMock')).vscodeMock({workspace:{findFiles:host.findFiles,workspaceFolders:[{uri:{fsPath:'C:/work'}}]}}));
import { ProjectExplorer, type XlideNode } from '../src/projectExplorer';
const BOOK='C:/work/App.vbp',OLD='C:/work/Old.vbp',NEW='C:/work/New.vbp';
const files=(path:string)=>[{scheme:'file',fsPath:path}];
const module=(name:string)=>[{name,type:'standard',folder:name+'Folder'}];
function deferred<T>(){let resolve!:(value:T)=>void,reject!:(reason:unknown)=>void;const promise=new Promise<T>((yes,no)=>{resolve=yes;reject=no;});return{promise,resolve,reject};}
let explorers:ProjectExplorer[]=[];
function create(call: (method: string) => Promise<unknown> = vi.fn(async()=>module('New'))){const explorer=new ProjectExplorer({call} as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);explorers.push(explorer);return explorer;}
const labels=(nodes:XlideNode[])=>nodes.map(node=>node.kind==='folder'?node.folder:node.moduleName??node.label);
beforeEach(()=>{host.findFiles.mockReset();host.findFiles.mockResolvedValue(files(BOOK));explorers=[];});
afterEach(()=>{for(const explorer of explorers)explorer.dispose();});

describe('root discovery ownership after refresh',()=>{
 it('retries an obsolete root listing without caching or registering its files',async()=>{
  const first=deferred<ReturnType<typeof files>>();host.findFiles.mockReturnValueOnce(first.promise).mockResolvedValue(files(NEW));
  const explorer=create(),old=explorer.getChildren();explorer.refresh();first.resolve(files(OLD));
  expect((await old).map(n=>n.filePath)).toEqual([NEW]);expect((await explorer.getChildren()).map(n=>n.filePath)).toEqual([NEW]);expect(explorer.getParent({kind:'module',label:'Unused',filePath:OLD,moduleName:'Unused'})).toBeUndefined();expect(host.findFiles).toHaveBeenCalledTimes(2);
 });
 it('keeps an already completed newer root when the old one arrives last',async()=>{
  const first=deferred<ReturnType<typeof files>>();host.findFiles.mockReturnValueOnce(first.promise).mockResolvedValue(files(NEW));
  const explorer=create(),old=explorer.getChildren();explorer.refresh();const current=await explorer.getChildren();first.resolve(files(OLD));
  expect(await old).toBe(current);expect(await explorer.getChildren()).toBe(current);expect(explorer.getParent({kind:'module',label:'Unused',filePath:OLD,moduleName:'Unused'})).toBeUndefined();expect(host.findFiles).toHaveBeenCalledTimes(2);
 });
 it('joins current discovery for all older and coalesced callers across two refreshes',async()=>{
  const a=deferred<ReturnType<typeof files>>(),b=deferred<ReturnType<typeof files>>(),c=deferred<ReturnType<typeof files>>();
  host.findFiles.mockReturnValueOnce(a.promise).mockReturnValueOnce(b.promise).mockReturnValueOnce(c.promise);
  const explorer=create(),first=explorer.getChildren(),alsoFirst=explorer.getChildren();explorer.refresh();const second=explorer.getChildren();explorer.refresh();const third=explorer.getChildren();
  a.resolve(files(OLD));b.resolve(files(BOOK));c.resolve(files(NEW));const results=await Promise.all([first,alsoFirst,second,third]);
  for(const result of results){expect(result.map(n=>n.filePath)).toEqual([NEW]);expect(result).toBe(results[3]);}expect(host.findFiles).toHaveBeenCalledTimes(3);
 });
 it('ignores an obsolete discovery failure while the current root succeeds',async()=>{
  const first=deferred<ReturnType<typeof files>>();host.findFiles.mockReturnValueOnce(first.promise).mockResolvedValue(files(NEW));
  const explorer=create(),old=explorer.getChildren();explorer.refresh();first.reject(new Error('obsolete discovery'));
  await expect(old).resolves.toMatchObject([{filePath:NEW}]);expect(host.findFiles).toHaveBeenCalledTimes(2);
 });
 it('still reports a current discovery failure and permits a retry',async()=>{
  host.findFiles.mockRejectedValueOnce(new Error('current discovery')).mockResolvedValue(files(NEW));const explorer=create();
  await expect(explorer.getChildren()).rejects.toThrow('current discovery');await expect(explorer.getChildren()).resolves.toMatchObject([{filePath:NEW}]);expect(host.findFiles).toHaveBeenCalledTimes(2);
 });
 it('keeps root ownership when a module folder invalidation changes only module data',async()=>{
  const first=deferred<ReturnType<typeof files>>();host.findFiles.mockReturnValueOnce(first.promise);const explorer=create(),pending=explorer.getChildren();
  explorer.setModuleFolder(BOOK,'M','Edited');explorer.forgetModuleFolder(BOOK,'M');first.resolve(files(BOOK));
  await expect(pending).resolves.toMatchObject([{filePath:BOOK}]);expect(host.findFiles).toHaveBeenCalledTimes(1);
 });
});

describe('project child rows after refresh',()=>{
 it.each(['tree','folders'] as const)('does not derive or retain obsolete %s rows',async(view)=>{
  const first=deferred<ReturnType<typeof module>>(),list=vi.fn().mockReturnValueOnce(first.promise).mockResolvedValue(module('New'));
  const call=vi.fn(async(method:string)=>method==='listModules'?list():{isPasswordProtected:false,isSigned:false});const explorer=create(call);explorer.setView(view);
  const [project]=await explorer.getChildren(),old=explorer.getChildren(project);explorer.refresh();const [current]=await explorer.getChildren();first.resolve(module('Old'));
  expect(labels(await old)).toEqual([view==='folders'?'NewFolder':'New']);expect(labels(await explorer.getChildren(current))).toEqual([view==='folders'?'NewFolder':'New']);
  expect(explorer.getModuleNode(BOOK,'Old')).toBeUndefined();expect(explorer.getFolderNode(BOOK,'OldFolder')).toBeUndefined();expect(list).toHaveBeenCalledTimes(2);
 });
 it('retries an obsolete empty-project metadata response',async()=>{
  const started=deferred<void>(),info=deferred<{hasVbaProject:boolean}>(),list=vi.fn().mockResolvedValueOnce([]).mockResolvedValue(module('New'));
  const call=vi.fn(async(method:string)=>{if(method==='listModules')return list();if(method==='hasVbaProject'){started.resolve();return info.promise;}return{isPasswordProtected:false,isSigned:false};});
  const explorer=create(call);explorer.setView('folders');const [project]=await explorer.getChildren(),old=explorer.getChildren(project);await started.promise;
  explorer.refresh();const [current]=await explorer.getChildren();info.resolve({hasVbaProject:false});
  expect(labels(await old)).toEqual(['NewFolder']);expect(labels(await explorer.getChildren(current))).toEqual(['NewFolder']);expect(list).toHaveBeenCalledTimes(2);
 });
 it('coalesces retries with a pending current project load',async()=>{
  const first=deferred<ReturnType<typeof module>>(),second=deferred<ReturnType<typeof module>>(),list=vi.fn().mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);
  const call=vi.fn(async(method:string)=>method==='listModules'?list():{isPasswordProtected:false,isSigned:false});const explorer=create(call);explorer.setView('folders');
  const [project]=await explorer.getChildren(),a=explorer.getChildren(project),b=explorer.getChildren(project);explorer.refresh();const [current]=await explorer.getChildren(),c=explorer.getChildren(current);
  first.resolve(module('Old'));second.resolve(module('New'));const results=await Promise.all([a,b,c]);for(const result of results)expect(labels(result)).toEqual(['NewFolder']);expect(list).toHaveBeenCalledTimes(2);
  const [folder]=await explorer.getChildren(current);expect((await explorer.getChildren(folder)).map(n=>n.moduleName)).toEqual(['New']);expect(explorer.getParent(folder)).toBe(current);
 });
});
