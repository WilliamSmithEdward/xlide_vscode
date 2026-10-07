import { describe, expect, it } from 'vitest';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';

function pair(source: string, exclusive: boolean): void {
 const tracker = createConditionalActivityTracker(parseModule(source));
 const span = (name: string) => { const start = source.indexOf(name); return {start, end:start + name.length}; };
 expect(tracker).toBeDefined();
 expect(tracker!.mutuallyExclusive(span('FirstValue'), span('SecondValue'))).toBe(exclusive);
 expect(tracker!.mutuallyExclusive(span('SecondValue'), span('FirstValue'))).toBe(exclusive);
 expect(tracker!.mutuallyExclusive(span('FirstValue'), span('FirstValue'))).toBe(false);
}
const open = (depth: number) => '#If UNKNOWN Then\n'.repeat(depth);
const close = (depth: number) => '#End If\n'.repeat(depth);
const first = 'Public FirstValue As Long\n', second = 'Public SecondValue As Long\n';

describe('conditional arm comparisons', () => {
 it.each([1, 100, 1000])('joins independent chains at depth %i', depth => pair(open(depth)+first+close(depth)+open(depth)+second+close(depth),false));
 it.each([1, 100, 1000])('separates deeply nested alternative arms at depth %i', depth => pair('#If OUTER Then\n'+open(depth)+first+close(depth)+'#Else\n'+open(depth)+second+close(depth)+'#End If\n',true));
 it.each([1, 100, 1000])('joins sibling chains within a shared outer arm at depth %i', depth => pair('#If OUTER Then\n'+open(depth)+first+close(depth)+open(depth)+second+close(depth)+'#End If\n',false));
 it.each([1, 100, 1000])('separates the innermost shared chain at depth %i', depth => pair(open(depth)+first+'#Else\n'+second+close(depth),true));
 it('joins an ancestor to its nested child', () => pair('#If OUTER Then\n'+first+open(100)+second+close(100)+'#End If\n',false));
 it('joins unconditional code with conditional code', () => pair(first+open(100)+second+close(100),false));
 it('keeps duplicate diagnostics for declarations in independent deep chains', () => {
  const block = open(1000)+'Public RepeatedValue As Long\n'+close(1000);
  const failures: unknown[]=[];
  const actual=analyzeModule(block+block,{onInternalError:e=>failures.push(e)});
  const shallow=analyzeModule(open(1)+'Public RepeatedValue As Long\n'+close(1)+open(1)+'Public RepeatedValue As Long\n'+close(1));
  expect(failures).toEqual([]);
  expect(actual.map(d=>({code:d.code,message:d.message,severity:d.severity}))).toEqual(shallow.map(d=>({code:d.code,message:d.message,severity:d.severity})));
  expect(actual.some(d=>d.message.includes('RepeatedValue'))).toBe(true);
 });
 it('does not invent duplicate diagnostics across deep alternative arms', () => {
  const source='#If OUTER Then\n'+open(1000)+'Public RepeatedValue As Long\n'+close(1000)+'#Else\n'+open(1000)+'Public RepeatedValue As Long\n'+close(1000)+'#End If\n';
  const failures: unknown[]=[];
  const actual=analyzeModule(source,{onInternalError:e=>failures.push(e)});
  expect(failures).toEqual([]);
  expect(actual.filter(d=>d.message.includes('RepeatedValue'))).toEqual([]);
 });
});
