import { expect, it, vi } from 'vitest';
import { applyVbaTextEdits, type VbaTextEdit } from '../src/analyzer/refactor/refactorTypes';

function legacy(source: string, edits: readonly VbaTextEdit[]): string {
 let out=source;for(const edit of [...edits].sort((a,b)=>b.span.start-a.span.start))out=out.slice(0,edit.span.start)+edit.newText+out.slice(edit.span.end);return out;
}
for(const count of [10,1000])it('slices source-sized work once at '+count+' replacements',()=>{
 const source='x'.repeat(count*100),edits=Object.freeze(Array.from({length:count},(_,i)=>Object.freeze({span:Object.freeze({start:i*100,end:i*100+1}),newText:'Y'})));
 const slice=String.prototype.slice;let slicedChars=0;
 const spy=vi.spyOn(String.prototype,'slice').mockImplementation(function(this: string,start?:number,end?:number){const part=slice.call(this,start,end);slicedChars+=part.length;return part;});
 let actual:string;try{actual=applyVbaTextEdits(source,edits);}finally{spy.mockRestore();}
 expect(actual!).toBe(('Y'+'x'.repeat(99)).repeat(count));expect(slicedChars).toBeLessThanOrEqual(source.length);
});
const source='A😀B\r\nC\rD\nE';
const shapes: VbaTextEdit[][]=[
 [],[{span:{start:0,end:0},newText:'before'}],
 [{span:{start:2,end:4},newText:'X'},{span:{start:0,end:1},newText:'Y'}],
 [{span:{start:1,end:1},newText:'first'},{span:{start:1,end:1},newText:'second'}],
 [{span:{start:1,end:3},newText:'range'},{span:{start:1,end:1},newText:'insert'}],
 [{span:{start:1,end:1},newText:'insert'},{span:{start:1,end:3},newText:'range'}],
 [{span:{start:4,end:7},newText:'X'},{span:{start:2,end:5},newText:'Y'}],
 [{span:{start:1,end:5},newText:'X'},{span:{start:2,end:3},newText:'Y'}],
 [{span:{start:2,end:3},newText:''},{span:{start:3,end:4},newText:''}],
 ...[-3,0.5,NaN,Infinity,-Infinity,source.length+10].map(start=>[{span:{start,end:start+1},newText:'X'},{span:{start:0,end:0},newText:'Y'}]),
 [{span:{start:4,end:1},newText:'X'},{span:{start:0,end:0},newText:'Y'}],
];
for(const [index,edits] of shapes.entries())it('preserves existing edit semantics for shape '+index,()=>{
 const before=edits.map(e=>({...e,span:{...e.span}}));expect(applyVbaTextEdits(source,edits)).toBe(legacy(source,edits));expect(edits).toEqual(before);
});
it('matches stable sequential slicing for 2000 independently generated edit sets',()=>{
 let seed=1234567;const next=()=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed;};
 for(let round=0;round<2000;round++){
  const text='x😀\r\ny'.repeat(next()%20),edits=Array.from({length:next()%15},()=>{const start=next()%(text.length+10)-3,end=next()%(text.length+10)-3;return {span:{start,end},newText:['','Q','\r\n','😀'][next()%4]};});
  expect(applyVbaTextEdits(text,edits)).toBe(legacy(text,edits));
 }
});
