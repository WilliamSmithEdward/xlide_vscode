import { beforeEach, expect, it, vi } from 'vitest';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
const work = vi.hoisted(() => ({ reads: 0 }));
vi.mock('../src/analyzer/symbols/editorModuleSymbols', async () => {
 const actual = await vi.importActual<typeof import('../src/analyzer/symbols/editorModuleSymbols')>('../src/analyzer/symbols/editorModuleSymbols');
 const views = new WeakMap<object, ReturnType<typeof actual.editorModuleSymbols>>();
 return { ...actual, editorModuleSymbols(...args: Parameters<typeof actual.editorModuleSymbols>) {
  const original = actual.editorModuleSymbols(...args);
  let view = views.get(original);
  if (!view) {
   const clones = new Map(original.all.map(symbol => [symbol, { ...symbol, ...(symbol.children ? { children: new Proxy(symbol.children, {get(target,key,receiver) {
    if(typeof key === 'string' && /^\d+$/.test(key)) work.reads++;
    return Reflect.get(target,key,receiver);
   }}) } : {}) }]));
   view = {...original, all:original.all.map(symbol=>clones.get(symbol)!), root:{...original.root,children:(original.root.children??[]).map(symbol=>clones.get(symbol)??symbol)}};
   views.set(original,view);
  }
  return view;
 }};
});
beforeEach(()=>{work.reads=0;});
for(const count of [10,1000]) for(const eol of ['\n','\r\n','\r']) it(`reuses the signature with ${count} locals and ${JSON.stringify(eol)}`,()=>{
 const source=['Private Function Target(ByVal Input As Long) As String',...Array.from({length:count},(_,i)=>`Dim Local${i} As Long`),'End Function'].join(eol);
 const start=source.indexOf('Target');
 const expected={signature:'Function Target(Input As Long) As String',details:['Declared in Module: Module','Visibility: Private'],span:{start,end:start+6}};
 expect(resolveHover(source,start+2)).toEqual(expected);
 work.reads=0;
 for(let round=0;round<10;round++) expect(resolveHover(source,start+2)).toEqual(expected);
 expect(work.reads).toBe(0);
});
it('preserves local formatting, fresh details/spans and source identity',()=>{
 const source='Public Function Target(Optional ByVal Input As Long = 2, ParamArray Rest() As Variant) As String\nTarget = "ok"\nEnd Function';
 const first=source.indexOf('Target'),second=source.indexOf('Target =');
 const expected={signature:'Function Target(Input As Long, Rest As Variant) As String',details:['Declared in Module: Module','Visibility: Public'],span:{start:first,end:first+6}};
 const initial=resolveHover(source,first)!;
 expect(initial).toEqual(expected);
 initial.details.push('caller mutation');initial.signature='changed';
 expect(resolveHover(source,second)).toEqual({...expected,span:{start:second,end:second+6}});
 expect(resolveHover(source,first,{moduleName:'Renamed'})).toEqual({...expected,details:['Declared in Module: Renamed','Visibility: Public']});
 const changed=source.replace('Input As Long','Input As Double').replace(') As String',') As Boolean');
 expect(resolveHover(changed,first)?.signature).toBe('Function Target(Input As Double, Rest As Variant) As Boolean');
 expect(resolveHover(source,first)).toEqual(expected);
});
it('keeps procedure kinds, visibility and missing types',()=>{
 for(const [header,closer,signature] of [
  ['Sub Plain(Input)','End Sub','Sub Plain(Input)'],
  ['Private Property Get Plain() As Long','End Property','Property Get Plain() As Long'],
  ['Public Property Let Plain(ByVal Input As Long)','End Property','Property Let Plain(Input As Long)'],
  ['Public Property Set Plain(ByVal Input As Object)','End Property','Property Set Plain(Input As Object)'],
 ]) {
  const source=header+'\n'+closer, start=source.indexOf('Plain');
  expect(resolveHover(source,start)?.signature).toBe(signature);
  expect(resolveHover(source,start)?.signature).toBe(signature);
 }
});
