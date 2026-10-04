import {afterEach,describe,expect,it,vi} from 'vitest';
import {conditionOperands,type ConditionForm} from '../src/analyzer/diagnostics/conditionOperands';
import {rawExpressionTokens} from '../src/analyzer/diagnostics/walker';
afterEach(()=>vi.restoreAllMocks());
describe('condition operand scan',()=>{
 it.each([10,100,1000])('bounds duplicate checks for %i independent operands',count=>{
  const tokens=rawExpressionTokens('If '+Array(count).fill('Flag').join(' And ')+' Then');
  let visits=0;const original=Array.prototype.some;
  vi.spyOn(Array.prototype,'some').mockImplementation(function(this:unknown[],callback,thisArg){return original.call(this,(value,index,array)=>{const hit=value as {index?:number;form?:string};if(hit&&typeof hit.index==='number'&&typeof hit.form==='string')visits++;return callback.call(thisArg,value,index,array);});});
  expect(conditionOperands(tokens)).toEqual(Array.from({length:count},(_,i)=>({index:1+i*2,form:'logical'})));
  expect(visits).toBeLessThanOrEqual(count);
 });
 const cases:Array<[string,Array<[number,ConditionForm]>]>=[
  ['If Flag Then',[[1,'condition']]],['ElseIf Flag Then',[[1,'condition']]],
  ['Do While Flag',[[2,'condition']]],['Do Until Flag',[[2,'condition']]],['Loop While Flag',[[2,'condition']]],['Loop Until Flag',[[2,'condition']]],['While Flag',[[1,'condition']]],['Select Case Flag',[[2,'select']]],
  ['x = Not Flag',[[3,'not']]],['x = Flag And Other',[[2,'logical'],[4,'logical']]],['x = IIf(Flag, 1, 2)',[[4,'iif']]],
  ['If [Flag] Then',[[1,'condition']]],['If Not Flag Then',[[2,'not']]],['If Flag Then x = Other And Third',[[1,'condition'],[5,'logical'],[7,'logical']]],
  ['x = obj.Flag And Other',[[6,'logical']]],['x = obj!Flag And Other',[[6,'logical']]],['x = Flag(1) And Other',[[7,'logical']]],
  ['x = obj.IIf(Flag, 1, 2)',[]],['If Flag + 1 Then',[]],['',[]],['If',[]],
 ];
 it.each(cases)('preserves operand forms for %s',(text,expected)=>{
  const tokens=rawExpressionTokens(text);for(const token of tokens)Object.freeze(token);Object.freeze(tokens);
  expect(conditionOperands(tokens)).toEqual(expected.map(([index,form])=>({index,form})));
 });
});
