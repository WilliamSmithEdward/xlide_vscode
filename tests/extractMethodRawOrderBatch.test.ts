import { describe, expect, it, vi } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
function fixture(names: string[], prefix = '', eol = '\n', padded = true) {
 const order = names.map((_, i) => names[(i * 37) % names.length]);
 const notes = Array.from({length:Math.ceil(order.length/3)},(_,i)=>"' "+order.slice(i*3,i*3+3).join(' ')).join(eol);
 const before = (padded ? Array(10000).fill("' "+'q'.repeat(90)).join(eol)+eol : '')+prefix+notes+eol+'Option Explicit'+eol+'Sub Main()'+eol+names.map(name=>'Dim '+name+' As Long').join(eol)+eol;
 const body=names.map(name=>name+' = 1').join(eol);
 return {source:before+body+eol+'End Sub'+eol,span:{start:before.length,end:before.length+body.length},name:'Work',names,eol,body};
}
function run(f: ReturnType<typeof fixture>) {
 const wanted=new Set(f.names), original=String.prototype.indexOf;
 let searches=0;
 const spy=vi.spyOn(String.prototype,'indexOf').mockImplementation(function(this:string,needle:string,position?:number){if(String(this)===f.source&&wanted.has(needle)&&position===undefined)searches++;return original.call(this,needle,position);});
 let result:ReturnType<typeof extractMethod>;
 try{result=extractMethod(f);}finally{spy.mockRestore();}
 if(!result!.ok)throw new Error(result!.reason);
 const sorted=[...f.names].sort((a,b)=>f.source.indexOf(a)-f.source.indexOf(b));
 expect(applyVbaTextEdits(f.source,result!.edits)).toContain('Private Sub Work()'+f.eol+sorted.map(name=>'Dim '+name+' As Long').join(f.eol)+f.eol+f.body);
 return searches;
}
describe('Extract Method raw local ordering batches',()=>{
 for(const n of [64,100,1000])it(`removes ${n} independent large-prefix searches`,()=>{
  const f=fixture(Array.from({length:n},(_,i)=>'local'+i.toString().padStart(4,'0')));
  expect(run(f)).toBe(0);
 });
 for(const eol of ['\n','\r\n','\r'])it(`preserves prefixes, internal overlaps, strings, comments and exact Unicode case under ${JSON.stringify(eol)}`,()=>{
  const names=['Beta','Alpha','arr','a','Café','Caf','Δ','Δelta','ababa','baba','aba','CaseName',...Array.from({length:52},(_,i)=>'local'+i.toString().padStart(4,'0'))];
  const f=fixture(names,"' Beta Alpha arr CAfÉ ababa casename"+eol+"' \"Δelta Café CaseName\""+eol,eol);
  expect(run(f)).toBe(0);
 });
 it('drops an already resolved common short name from active matching',()=>{
  const names=['a',...Array.from({length:99},(_,i)=>'local'+i.toString().padStart(4,'0'))];
  const f=fixture(names);f.source=f.source.replace(/q/g,'a');
  expect(run(f)).toBe(0);
 });
 for(const n of [1,10,63])it(`retains direct ordering for ${n} locals even with a large prefix`,()=>expect(run(fixture(Array.from({length:n},(_,i)=>'local'+i)))).toBe(n===1?0:n));
 it('retains the direct path for a small source workload',()=>expect(run(fixture(Array.from({length:100},(_,i)=>'local'+i),'','\n',false))).toBe(100));
 it('bounds dictionary allocation for long names',()=>expect(run(fixture(Array.from({length:400},(_,i)=>'local'+'x'.repeat(200)+i)))).toBe(400));
 for(const ErrorType of [SyntaxError,RangeError])it(`preserves results if the engine rejects the dictionary with ${ErrorType.name}`,()=>{
  const names=Array.from({length:64},(_,i)=>'local'+i.toString().padStart(4,'0')),f=fixture(names),Original=RegExp;
  vi.stubGlobal('RegExp',new Proxy(Original,{construct(target,args){if(String(args[0]).includes('local0000')&&args[1]==='g')throw new ErrorType('test engine limit');return Reflect.construct(target,args);}}));
  try{expect(run(f)).toBe(64);}finally{vi.unstubAllGlobals();}
 });
});
it('does not use unrelated suffix size to batch an early small ordering query',()=>{
 const names=Array.from({length:100},(_,i)=>'local'+i),f=fixture(names,'','\n',false);
 f.source+=Array(10000).fill("' "+'q'.repeat(90)).join('\n')+'\n';
 expect(run(f)).toBe(100);
});
