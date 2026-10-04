import { describe, expect, it, vi } from 'vitest';
import { conditionalActivityAtOffset, createConditionalActivityTracker, indexConditionalCompilation, nullConditionDirectives } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';

const at = (source: string) => {const start=source.indexOf('Public Taken');return {start,end:start+6};};
describe('conditional compiler lookup reuse',()=>{
 it.each(['tracker','index','offset','nulls'] as const)('builds compiler lookups once per %s replay',kind=>{
  const source=Array.from({length:400},(_,i)=>'#Const STEP'+i+' = '+(i===0?'1':'STEP'+(i-1)+' + 1')+'\n').join('')+'#If STEP399 = 400 Then\nPublic Taken As Long\n#End If\n';
  const module=parseModule(source),original=Map.prototype.set;let compilerSets=0;
  const spy=vi.spyOn(Map.prototype,'set').mockImplementation(function(this:Map<unknown,unknown>,key:unknown,value:unknown){if(key==='vba7')compilerSets++;return original.call(this,key,value);});
  try {if(kind==='tracker')expect(createConditionalActivityTracker(module)!.activityForSpan(at(source))).toBe('active');else if(kind==='index')expect(indexConditionalCompilation(module).constants[399].value).toBe(400);else if(kind==='offset')expect(conditionalActivityAtOffset(module,at(source).start)).toBe('active');else expect(nullConditionDirectives(module)).toEqual([]);expect(compilerSets).toBe(1);}finally{spy.mockRestore();}
 });
 it('uses evolving module constants before compiler constants',()=>{
  const source='#Const FLAG = 0\n#If FLAG Then\nPublic Off As Long\n#End If\n#Const FLAG = 1\n#If FLAG Then\nPublic Taken As Long\n#End If\n';const tracker=createConditionalActivityTracker(parseModule(source),{compilerConstants:{FLAG:9},projectConstants:{FLAG:7}})!;
  expect(tracker.activityForSpan({start:source.indexOf('Public Off'),end:source.indexOf('Public Off')+6})).toBe('inactive');expect(tracker.activityForSpan(at(source))).toBe('active');
 });
 it('preserves distinct index and activity missing-name policies',()=>{
  const source='#Const FLAG = MISSING + 1\n#If FLAG = 1 Then\nPublic Taken As Long\n#End If\n',module=parseModule(source);
  expect(indexConditionalCompilation(module).constants[0].value).toBe(1);expect(createConditionalActivityTracker(module)!.activityForSpan(at(source))).toBe('unknown');expect(createConditionalActivityTracker(module,{projectConstants:{}})!.activityForSpan(at(source))).toBe('active');
 });
 it('rebuilds compiler lookups for fresh query environments',()=>{
  const source='#If FLAG Then\nPublic Taken As Long\n#End If\n',module=parseModule(source),env={compilerConstants:{FLAG:0}};
  expect(createConditionalActivityTracker(module,env)!.activityForSpan(at(source))).toBe('inactive');env.compilerConstants.FLAG=1;expect(createConditionalActivityTracker(module,env)!.activityForSpan(at(source))).toBe('active');expect(env.compilerConstants.FLAG).toBe(1);
 });
 it('does not build compiler lookups when no expression needs evaluation',()=>{
  const module=parseModule('#Const FLAG\n');const original=Map.prototype.set;let compilerSets=0;const spy=vi.spyOn(Map.prototype,'set').mockImplementation(function(this:Map<unknown,unknown>,key:unknown,value:unknown){if(key==='vba7')compilerSets++;return original.call(this,key,value);});
  try{indexConditionalCompilation(module);createConditionalActivityTracker(module);nullConditionDirectives(module);conditionalActivityAtOffset(module,100);expect(compilerSets).toBe(0);}finally{spy.mockRestore();}
 });
});
