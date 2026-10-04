import {afterEach,expect,it,vi} from 'vitest';
import {introduceParameter} from '../src/analyzer/refactor/introduceParameter';
afterEach(()=>vi.restoreAllMocks());
for(const count of [10,100,1000])it('bounds stranded-name membership scans for '+count+' names',()=>{
 const names=Array.from({length:count},(_,i)=>'v'+i),chunks:string[]=[];for(let i=0;i<count;i+=100)chunks.push(names.slice(i,i+100).join(' + '));
 const source=[...names.map(n=>'Private '+n+' As Long'),'Sub Report()','Dim limit As Long','limit = '+chunks.join(' + _\n'),'Debug.Print limit','End Sub',''].join('\n');
 let searchedSlots=0;const original=Array.prototype.includes;
 vi.spyOn(Array.prototype,'includes').mockImplementation(function(this:unknown[],value,fromIndex){if(typeof value==='string'&&/^v\d+$/.test(value)&&this.every(v=>typeof v==='string'&&/^v\d+$/.test(v)))searchedSlots+=this.length;return original.call(this,value,fromIndex);});
 const result=introduceParameter({source,offset:source.indexOf('limit As'),moduleName:'M'});
 expect(result).toEqual({ok:false,reason:"The value '"+names.join(' + ')+"' names "+names.map(n=>"'"+n+"'").join(', ')+', which a caller in another module cannot see.'});
 expect(searchedSlots).toBeLessThanOrEqual(count);
});
it.each([
 ['value + value',"'value'"],['value + VALUE + value',"'value', 'VALUE'"],['"value" + value',"'value'"],['absent + value',"'value'"],['value + other + value',"'value', 'other'"],
])('preserves first occurrence order and exact spelling for %s',(value,names)=>{
 const source=['Private value As Long','Private other As Long','Sub Report()','Dim limit As Long','limit = '+value,'Debug.Print limit','End Sub',''].join('\n');
 expect(introduceParameter({source,offset:source.indexOf('limit As'),moduleName:'M'})).toEqual({ok:false,reason:"The value '"+value+"' names "+names+', which a caller in another module cannot see.'});
});
