import { describe, expect, it, vi } from 'vitest';
import { operatorYieldsNull } from '../src/analyzer/diagnostics/nullOperators';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { tokenWord } from '../src/analyzer/lexer/tokenHelpers';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
const tokens = (expression: string) => tokenize('x='+expression).slice(2);
const holdsNull = (token: ReturnType<typeof tokens>[number]) => tokenWord(token)==='null' || tokenWord(token)==='held';

describe('Null inference stack and token work',()=>{
 it.each([
  ['Not '.repeat(20000)+'Null',true],
  ['- '.repeat(20000)+'held',true],
  ['Abs('.repeat(2000)+'Null'+')'.repeat(2000),true],
  ['1 + ('.repeat(1000)+'Null'+')'.repeat(1000),true],
  ['Null And ('.repeat(1000)+'Null'+')'.repeat(1000),true],
 ] as const)('evaluates deep expression family %#', (expression,expected)=>{
  const input=Object.freeze(tokens(expression).map(t=>Object.freeze(t)));expect(operatorYieldsNull(input,holdsNull)).toBe(expected);
 });
 it.each([100,1000])('copies linear token work through %i unary operators',count=>{
  const input=Object.freeze(tokens('Not '.repeat(count)+'Null'));const original=Array.prototype.slice;let copied=0;
  const spy=vi.spyOn(Array.prototype,'slice').mockImplementation(function(this: unknown[],start?:number,end?:number){const result=original.call(this,start,end);copied+=result.length;return result;});
  let actual: boolean;try {actual=operatorYieldsNull(input,holdsNull);}finally{spy.mockRestore();}
  expect(actual!).toBe(true);expect(copied).toBeLessThanOrEqual(input.length);
 });
 it.each([100,1000])('reads linear token work through %i nested arithmetic operands',depth=>{
  const original=tokens('1 + ('.repeat(depth)+'Null'+')'.repeat(depth));let reads=0;
  const input=Object.freeze(original.map(t=>Object.freeze({...t,get rawText(){reads++;return t.rawText;}})));
  expect(operatorYieldsNull(input,holdsNull)).toBe(true);expect(reads).toBeLessThanOrEqual(input.length*40);
 });
 it.each(['assignment','argument'] as const)('keeps the scalar %s diagnostic with 10000 Not operators',kind=>{
  const expression='Not '.repeat(10000)+'Null';
  const source=kind==='assignment'?'Option Explicit\nSub Go()\nDim value As Long\nvalue = '+expression+'\nDebug.Print value\nEnd Sub\n':'Option Explicit\nSub Take(ByVal value As Long)\nDebug.Print value\nEnd Sub\nSub Go()\nTake ('+expression+')\nEnd Sub\n';
  const failures: unknown[]=[];const diagnostics=analyzeModule(source,{onInternalError:e=>failures.push(e)});
  expect(failures).toEqual([]);
  const diagnostic=diagnostics.find(d=>d.code===(kind==='assignment'?'assignment-type-mismatch':'argument-type-mismatch'));
  expect(diagnostic).toBeDefined();expect(diagnostic!.message).toContain('Null');
  const start=source.indexOf(expression);expect(diagnostic!.span).toEqual({start,end:start+expression.length});
 });
 it.each(['assignment','argument'] as const)('restores the continued scalar %s diagnostic within physical limits',kind=>{
  const expression=Array(20).fill('Not '.repeat(250).trimEnd()).join(' _\n')+' Null';
  const source=kind==='assignment'?'Option Explicit\nSub Go()\nDim value As Long\nvalue = '+expression+'\nDebug.Print value\nEnd Sub\n':'Option Explicit\nSub Take(ByVal value As Long)\nDebug.Print value\nEnd Sub\nSub Go()\nTake ('+expression+')\nEnd Sub\n';
  expect(Math.max(...source.split('\n').map(line=>line.length))).toBeLessThanOrEqual(1023);
  const failures:unknown[]=[];const diagnostics=analyzeModule(source,{onInternalError:e=>failures.push(e)});
  expect(failures).toEqual([]);expect(diagnostics).toHaveLength(1);
  expect(diagnostics[0].code).toBe(kind==='assignment'?'assignment-type-mismatch':'argument-type-mismatch');
  const start=source.indexOf(expression);expect(diagnostics[0].span).toEqual({start,end:start+expression.length});
 });
 it.each([
  ['Null + 1',true], ['1 + Null',true], ['Null & "x"',false], ['Null And 0',false],
  ['Null And 1',true], ['40000 Or Null',false], ['Null Or 0',true], ['Null Or True',false],
  ['Null Imp 12',false], ['False Imp Null',false], ['Null Imp False',true], ['Null Imp True',false],
  ['Null Imp Null',true], ['Not Null',true], ['Abs(Null)',true], ['Abs(Null + 1)',true],
  ['(Null)',true], ['Null Xor 1',true], ['Null Eqv 1',true], ['Null And Null Or Null',true],
  ['Null And 1 Or Null',false], ['"Null"',false], ['1 + 2',false], ['Null +',false],
  ['',false], ['Abs()',false], ['Not Missing',false], ['-held',true],
 ] as const)('retains Null inference: %s',(expression,expected)=>expect(operatorYieldsNull(tokens(expression),holdsNull)).toBe(expected));
 it.each([
  ['Null + Later',['null']], ['1 + Null + Later',['1','null']],
  ['Null And False',['null','false']], ['Null And Missing Or Later',['null','missing']],
  ['Not Abs(Null)',['null']], ['Null And (1 + Later)',['null','1','later']],
 ] as const)('retains callback order and short-circuiting: %s',(expression,expected)=>{
  const calls:string[]=[];operatorYieldsNull(tokens(expression),t=>{calls.push(tokenWord(t));return holdsNull(t);});expect(calls).toEqual(expected);
 });
 it('does not swallow callback errors',()=>{
  const failure=new Error('callback failed');expect(()=>operatorYieldsNull(tokens('1 + Null'),()=>{throw failure;})).toThrow(failure);
 });
});
