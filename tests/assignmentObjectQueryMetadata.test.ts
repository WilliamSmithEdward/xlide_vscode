import {expect,it} from 'vitest';
import {analyzeModule} from '../src/analyzer/diagnostics/analyzeModule';
import type {VbaProjectClassMembers} from '../src/analyzer/symbols/symbolModel';
type Mode='variable'|'held'|'property';
const modes:Mode[]=['variable','held','property'];
function surface(name:string,implemented:string[]=[]):VbaProjectClassMembers{return {name,moduleName:name,kind:'class',exhaustive:true,members:[],implements:implemented};}
function classesFor(mode:Mode,implemented:string[]=[]){const classes=[surface('Class1'),surface('Class2',implemented)];if(mode==='property'){const holder=surface('Holder');holder.members=[{name:'Item',kind:'property',moduleName:'Holder',returns:'Class1',writeType:'Class1',writable:true,setAccessor:true}];classes.push(holder);}return classes;}
function sourceFor(mode:Mode,count:number,procedures=false){return ['Option Explicit',...Array.from({length:procedures?count:1},(_,i)=>['Sub Go'+i+'(ByVal actual As '+(mode==='held'?'Object':'Class2')+', ByVal target As Class1'+(mode==='property'?', ByVal holder As Holder':'')+')',...(mode==='held'?['Set actual = New Class2']:[]),...Array.from({length:procedures?1:count},()=>mode==='property'?'Set holder.Item = actual':'Set target = actual'),'End Sub'].join('\n')),''].join('\n');}
function run(source:string,classes:VbaProjectClassMembers[]){const errors:string[]=[];const out=analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});expect(errors).toEqual([]);return out;}
for(const mode of modes){
 it.each([10,100,1000])('bounds all consulted assignment type/interface work at %i statements: '+mode,count=>{
  let reads=0,names=0;const implemented=new Proxy(Array.from({length:count},(_,i)=>i===count-1?'Class1':'Other'+i),{get(t,k,r){if(typeof k==='string'&&/^(0|[1-9]\d*)$/.test(k)){reads++;}return Reflect.get(t,k,r);}});
  const classes=classesFor(mode,implemented);classes.push(...Array.from({length:count},(_,i)=>surface('Extra'+i)));for(const type of classes){const name=type.name;Object.defineProperty(type,'name',{get(){names++;return name;}});}
  expect(run(sourceFor(mode,count),classes)).toEqual([]);expect(reads).toBeLessThanOrEqual(count*3);expect(names).toBeLessThanOrEqual(count*25+100);
 });
 it('shares the '+mode+' membership lookup across procedures',()=>{
  let reads=0;const names=new Proxy(Array.from({length:100},(_,i)=>i===99?'Class1':'Other'+i),{get(t,k,r){if(typeof k==='string'&&/^(0|[1-9]\d*)$/.test(k)){reads++;}return Reflect.get(t,k,r);}});
  expect(run(sourceFor(mode,100,true),classesFor(mode,names))).toEqual([]);expect(reads).toBeLessThanOrEqual(300);
 });
 it('preserves the exact '+mode+' incompatible assignment',()=>{
  const source=sourceFor(mode,1),out=run(source,classesFor(mode)),start=source.lastIndexOf('actual');expect(out).toHaveLength(1);
  expect(out[0].code).toBe('assignment-object-type-mismatch');expect(out[0].span).toEqual({start,end:start+6});
  expect(out[0].message).toBe("Object assignment to '"+(mode==='property'?'holder.Item':'target')+"' expects Class1, but got "+(mode==='held'?"'actual', which holds a Class2 here":"actual As Class2")+". This object type is not compatible with Class1. This will raise Run-time error '13': Type mismatch.");
 });
 it('refreshes '+mode+' interface facts in the next module query',()=>{const source=sourceFor(mode,1),classes=classesFor(mode);expect(run(source,classes)).toHaveLength(1);classes[1].implements=['cLaSs1'];expect(run(source,classes)).toEqual([]);classes[1].implements=[];expect(run(source,classes)).toHaveLength(1);});
}
