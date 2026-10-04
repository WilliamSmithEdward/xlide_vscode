import { afterEach, expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import * as facts from '../src/analyzer/diagnostics/openedFileNumbers';
afterEach(()=>vi.restoreAllMocks());
const source=(number:string,name='P')=>'Sub '+name+'()\nOpen "file" For Input As #'+number+'\nEnd Sub\n';
for(const count of [2,20,100])it('rescans only the edited module in a '+count+' module project',()=>{
 const index=new ProjectIndex();for(let i=0;i<count;i++)index.setModule({moduleName:'M'+i,moduleKind:'standard',source:source(String(i+1),'P'+i)});
 const spy=vi.spyOn(facts,'openedFileNumbersIn');
 expect(index.openedFileNumbers()).toEqual({any:false,numbers:new Set(Array.from({length:count},(_,i)=>i+1))});expect(spy).toHaveBeenCalledTimes(count);spy.mockClear();
 for(let edit=0;edit<3;edit++){const number=count+10+edit;index.setModule({moduleName:'m0',moduleKind:'standard',source:source(String(number),'Changed')});expect(index.openedFileNumbers()).toEqual({any:false,numbers:new Set([number,...Array.from({length:count-1},(_,i)=>i+2)])});}
 expect(spy).toHaveBeenCalledTimes(3);
});
it('retains untouched facts through removals, additions and unknown-number replacement',()=>{
 const index=new ProjectIndex();index.setModule({moduleName:'A',moduleKind:'standard',source:source('1')});index.setModule({moduleName:'B',moduleKind:'standard',source:source('2')});index.openedFileNumbers();const spy=vi.spyOn(facts,'openedFileNumbersIn');
 index.removeModule('b');expect(index.openedFileNumbers()).toEqual({any:false,numbers:new Set([1])});expect(spy).not.toHaveBeenCalled();
 index.setModule({moduleName:'B',moduleKind:'standard',source:source('f')});expect(index.openedFileNumbers()).toEqual({any:true,numbers:new Set([1])});expect(spy).toHaveBeenCalledTimes(1);
 index.setModule({moduleName:'b',moduleKind:'standard',source:source('3')});expect(index.openedFileNumbers()).toEqual({any:false,numbers:new Set([1,3])});expect(spy).toHaveBeenCalledTimes(2);
 index.removeModule('A');index.removeModule('B');expect(index.openedFileNumbers()).toEqual({any:false,numbers:new Set()});expect(spy).toHaveBeenCalledTimes(2);
});
it('keeps initial scans lazy and stable-revision queries memoized',()=>{const index=new ProjectIndex();const spy=vi.spyOn(facts,'openedFileNumbersIn');index.setModule({moduleName:'A',moduleKind:'standard',source:source('1')});expect(spy).not.toHaveBeenCalled();const first=index.openedFileNumbers();expect(index.openedFileNumbers()).toBe(first);expect(spy).toHaveBeenCalledTimes(1);});
