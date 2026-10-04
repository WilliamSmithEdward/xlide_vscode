import { afterEach, expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import * as facts from '../src/analyzer/symbols/sheetChanges';
afterEach(()=>vi.restoreAllMocks());
const source=(name:string,adds=false,computed=false)=>'Sub P()\nws.Name = "'+name+'"\n'+(adds?'Worksheets.Add\n':'')+(computed?'ws.Name = value\n':'')+'End Sub\n';
for(const count of [2,20,100])it('rescans only the edited source in a '+count+' module sheet query',()=>{
 const index=new ProjectIndex();for(let i=0;i<count;i++)index.setModule({moduleName:'M'+i,moduleKind:'standard',source:source('N'+i,i===0)});const spy=vi.spyOn(facts,'sheetChangesIn');
 expect(index.sheetChanges()).toEqual({addsSheets:true,namesAssigned:new Set(Array.from({length:count},(_,i)=>'n'+i)),assignsComputedName:false});expect(spy).toHaveBeenCalledTimes(count);spy.mockClear();
 for(let edit=0;edit<3;edit++){index.setModule({moduleName:'m0',moduleKind:'standard',source:source('Changed'+edit)});const result=index.sheetChanges();expect(result).toEqual({addsSheets:false,namesAssigned:new Set(['changed'+edit,...Array.from({length:count-1},(_,i)=>'n'+(i+1))]),assignsComputedName:false});expect([...result.namesAssigned]).toEqual(['changed'+edit,...Array.from({length:count-1},(_,i)=>'n'+(i+1))]);}
 expect(spy).toHaveBeenCalledTimes(3);
});
it('drops removed additions/computed names without scanning survivors and rebuilds replaced roles',()=>{
 const index=new ProjectIndex();index.setModule({moduleName:'A',moduleKind:'standard',source:source('Alpha')});index.setModule({moduleName:'B',moduleKind:'class',source:source('Beta',true,true)});index.sheetChanges();const spy=vi.spyOn(facts,'sheetChangesIn');
 index.removeModule('b');expect(index.sheetChanges()).toEqual({addsSheets:false,namesAssigned:new Set(['alpha']),assignsComputedName:false});expect(spy).not.toHaveBeenCalled();
 index.setModule({moduleName:'B',moduleKind:'document',source:source('Beta',true,true)});expect(index.sheetChanges()).toEqual({addsSheets:true,namesAssigned:new Set(['alpha','beta']),assignsComputedName:true});expect(spy).toHaveBeenCalledTimes(1);
 index.setModule({moduleName:'b',moduleKind:'userform',source:source('Gamma')});expect(index.sheetChanges()).toEqual({addsSheets:false,namesAssigned:new Set(['alpha','gamma']),assignsComputedName:false});expect(spy).toHaveBeenCalledTimes(2);
 index.removeModule('A');index.removeModule('B');expect(index.sheetChanges()).toEqual({addsSheets:false,namesAssigned:new Set(),assignsComputedName:false});expect(spy).toHaveBeenCalledTimes(2);
});
it('keeps sheet scans lazy and stable-revision results memoized',()=>{const index=new ProjectIndex();const spy=vi.spyOn(facts,'sheetChangesIn');index.setModule({moduleName:'A',moduleKind:'standard',source:source('Alpha')});expect(spy).not.toHaveBeenCalled();const first=index.sheetChanges();expect(index.sheetChanges()).toBe(first);expect(spy).toHaveBeenCalledTimes(1);});
