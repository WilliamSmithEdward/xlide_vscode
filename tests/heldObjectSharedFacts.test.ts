import { afterEach, expect, it, vi } from 'vitest';
import { heldObjectsAt } from '../src/analyzer/diagnostics/heldObjects';
import * as flow from '../src/analyzer/diagnostics/dataflow';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
function fixture(body:string){const source='Sub Work()\n'+body+'\nEnd Sub';const module=parseModule(source);const proc=module.members.find(m=>m.kind==='Procedure')!;if(proc.kind!=='Procedure')throw Error('procedure');const symbols=buildModuleSymbols('M','standard',source,{parsedModule:module});return{source,module,proc,symbols,read:proc.body[proc.body.length-1]};}
afterEach(()=>vi.restoreAllMocks());
for(const count of [2,10,100])it('shares the complete default body walk across '+count+' equivalent consumers',()=>{
 const f=fixture('Dim c As New Collection\nc.Add 1\nDebug.Print c.Count');const spy=vi.spyOn(flow,'walkEnteringBlocks');
 for(let i=0;i<count;i++){const at=heldObjectsAt(f.source,f.proc,f.symbols,undefined);expect([...at(f.read).classes]).toEqual([['c','Collection']]);expect([...at(f.read).items]).toEqual([['c',['(value)']]]);}
 expect(spy).toHaveBeenCalledTimes(1);
});
it('keeps changed source separate with reused procedure and symbols',()=>{
 const f=fixture('Dim o As Object\nSet o = New A\nDebug.Print o.Value');
 expect(heldObjectsAt(f.source,f.proc,f.symbols,undefined)(f.read).classes.get('o')).toBe('A');
 expect(heldObjectsAt(f.source.replace('New A','New B'),f.proc,f.symbols,undefined)(f.read).classes.get('o')).toBe('B');
 expect(heldObjectsAt(f.source,f.proc,f.symbols,undefined)(f.read).classes.get('o')).toBe('A');
});
it('keys facts by the symbol snapshot',()=>{
 const f=fixture('Dim c As New Collection\nDebug.Print c.Count');const other=buildModuleSymbols('Other','standard',f.source,{parsedModule:f.module});const spy=vi.spyOn(flow,'walkEnteringBlocks');
 heldObjectsAt(f.source,f.proc,f.symbols,undefined);heldObjectsAt(f.source,f.proc,other,undefined);expect(spy).toHaveBeenCalledTimes(2);
});
it('keeps conditional activity snapshots separate and rebuilds a replaced entry',()=>{
 const f=fixture('Dim o As New A\n#If VBA7 Then\nSet o = New B\n#End If\nDebug.Print o.Value');const off=createConditionalActivityTracker(f.module,{compilerConstants:{VBA7:false}}),on=createConditionalActivityTracker(f.module,{compilerConstants:{VBA7:true}});const spy=vi.spyOn(flow,'walkEnteringBlocks');
 for(const [activity,expected] of [[off,'A'],[on,'B'],[off,'A']] as const){for(let i=0;i<2;i++)expect(heldObjectsAt(f.source,f.proc,f.symbols,activity)(f.read).classes.get('o')).toBe(expected);}
 expect(spy).toHaveBeenCalledTimes(3);
});
it('does not cache callback-driven facts, even when the callback identity is unchanged',()=>{
 const f=fixture('Dim o As Object\nSet o = MakeThing()\nDebug.Print o.Value');let type='A';const resolver=vi.fn(()=>type);
 expect(heldObjectsAt(f.source,f.proc,f.symbols,undefined,resolver)(f.read).classes.get('o')).toBe('A');type='B';
 expect(heldObjectsAt(f.source,f.proc,f.symbols,undefined,resolver)(f.read).classes.get('o')).toBe('B');expect(resolver).toHaveBeenCalledTimes(2);
 expect(heldObjectsAt(f.source,f.proc,f.symbols,undefined)(f.read).classes.get('o')).toBeUndefined();
});
