import { describe, expect, it, vi } from 'vitest';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { conditionalActivityAtOffset, createConditionalActivityTracker, indexConditionalCompilation, nullConditionDirectives } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';

const spanAt = (source: string,name: string) => {const start=source.indexOf(name);return {start,end:start+name.length};};

describe('conditional inactive-arm evaluation gates', () => {
 it.each([10,1000])('skips %i settled ElseIf expressions during tracker replay',count=>{
  const source='#If True Then\nPublic Taken As Long\n'+'#ElseIf Not Missing Then\n'.repeat(count)+'Public Skipped As Long\n#End If\n';
  const module=parseModule(source),spy=vi.spyOn(lexer,'tokenize');
  try {const tracker=createConditionalActivityTracker(module)!;expect(tracker.activityForSpan(spanAt(source,'Taken'))).toBe('active');expect(tracker.activityForSpan(spanAt(source,'Skipped'))).toBe('inactive');expect(spy).toHaveBeenCalledTimes(1);} finally {spy.mockRestore();}
 });
 it.each([10,1000])('skips %i nested conditions under an inactive parent',count=>{
  const source='#If False Then\n'+'#If Not Missing Then\n'.repeat(count)+'Public Skipped As Long\n'+'#End If\n'.repeat(count)+'#Else\nPublic Taken As Long\n#End If\n';
  const module=parseModule(source),spy=vi.spyOn(lexer,'tokenize');
  try {const tracker=createConditionalActivityTracker(module)!;expect(tracker.activityForSpan(spanAt(source,'Skipped'))).toBe('inactive');expect(tracker.activityForSpan(spanAt(source,'Taken'))).toBe('active');expect(spy).toHaveBeenCalledTimes(1);} finally {spy.mockRestore();}
 });
 it('gates offset replay and null-condition scanning through the shared replay path',()=>{
  const source='#If True Then\nPublic Taken As Long\n'+'#ElseIf Null Then\n'.repeat(100)+'Public Skipped As Long\n#End If\n';
  const module=parseModule(source),spy=vi.spyOn(lexer,'tokenize');
  try {expect(conditionalActivityAtOffset(module,spanAt(source,'Skipped').start)).toBe('inactive');expect(spy).toHaveBeenCalledTimes(1);spy.mockClear();expect(nullConditionDirectives(module)).toEqual([]);expect(spy).toHaveBeenCalledTimes(2);} finally {spy.mockRestore();}
 });
 it('continues replaying constants in skipped arms and inactive nested chains',()=>{
  const source='#If True Then\n#ElseIf Null Then\n#Const FLAG = 1\n#If False Then\n#Const FLAG = FLAG + 1\n#ElseIf Null Then\n#Const FLAG = FLAG + 1\n#End If\n#End If\n#If FLAG = 3 Then\nPublic Taken As Long\n#Else\nPublic Skipped As Long\n#End If\n';
  const module=parseModule(source),tracker=createConditionalActivityTracker(module)!;
  expect(tracker.activityForSpan(spanAt(source,'Taken'))).toBe('active');expect(tracker.activityForSpan(spanAt(source,'Skipped'))).toBe('inactive');expect(indexConditionalCompilation(module).constants.map(c=>c.value)).toEqual([1,2,3]);
 });
 it('evaluates ElseIf after a false arm and discovers active later arms',()=>{
  const source='#If False Then\nPublic Skipped As Long\n#ElseIf True Then\nPublic Taken As Long\n#Else\nPublic Later As Long\n#End If\n';const tracker=createConditionalActivityTracker(parseModule(source))!;
  expect(tracker.activityForSpan(spanAt(source,'Skipped'))).toBe('inactive');expect(tracker.activityForSpan(spanAt(source,'Taken'))).toBe('active');expect(tracker.activityForSpan(spanAt(source,'Later'))).toBe('inactive');
 });
 it('retains uncertain-arm evaluation and excludes later arms after a possible true arm',()=>{
  const source='#If Missing Then\nPublic First As Long\n#ElseIf False Then\nPublic Skipped As Long\n#ElseIf True Then\nPublic Maybe As Long\n#ElseIf Null Then\nPublic Later As Long\n#Else\nPublic Last As Long\n#End If\n';const tracker=createConditionalActivityTracker(parseModule(source))!;
  expect(tracker.activityForSpan(spanAt(source,'First'))).toBe('unknown');expect(tracker.activityForSpan(spanAt(source,'Skipped'))).toBe('inactive');expect(tracker.activityForSpan(spanAt(source,'Maybe'))).toBe('unknown');expect(tracker.activityForSpan(spanAt(source,'Later'))).toBe('inactive');expect(tracker.activityForSpan(spanAt(source,'Last'))).toBe('inactive');
 });
 it('still reports null conditions that are definitely evaluated',()=>{
  const source='#If False Then\n#ElseIf Null Then\nPublic Maybe As Long\n#End If\n#If Null Then\nPublic Other As Long\n#End If\n';expect(nullConditionDirectives(parseModule(source)).map(d=>d.directiveKind)).toEqual(['ElseIf','If']);
 });
 it('retains fresh project constants and does not mutate caller environments',()=>{
  const source='#If FLAG Then\nPublic First As Long\n#ElseIf True Then\nPublic Second As Long\n#End If\n';const module=parseModule(source),env={projectConstants:{FLAG:0}};
  expect(createConditionalActivityTracker(module,env)!.activityForSpan(spanAt(source,'Second'))).toBe('active');env.projectConstants.FLAG=1;expect(createConditionalActivityTracker(module,env)!.activityForSpan(spanAt(source,'Second'))).toBe('inactive');expect(env.projectConstants.FLAG).toBe(1);
 });
});
