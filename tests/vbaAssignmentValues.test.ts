import { afterEach, describe, expect, it, vi } from 'vitest';
import { resolveAssignmentValueCompletion } from '../src/analyzer/completion/assignmentValueCompletion';
import { assignmentValueTriggerMayComplete, spaceTriggerMayComplete } from '../src/analyzer/completion/cursorContext';
import { resolveArgumentValueCompletion } from '../src/analyzer/completion/argumentValueCompletion';

function accepted(line: string, prelude = '') {
	const source = `${prelude}\nSub Demo()\n${line}\nEnd Sub`;
	return resolveAssignmentValueCompletion(source, source.indexOf(line) + line.length);
}

describe('assignment value completion', () => {
	it('respects the declared enum type of a custom Color property', () => {
		const source = 'Sub T()\nDim item As Widget\nitem.Color = ';
		const result = resolveAssignmentValueCompletion(source, source.length, { projectClassMembers: [
			{ name: 'Widget', moduleName: 'Widget', kind: 'class', members: [{ name: 'Color', kind: 'property', declaredType: 'Shade', writable: true }] },
			{ name: 'Shade', moduleName: 'Types', kind: 'enum', members: [{ name: 'Light', kind: 'property' }, { name: 'Dark', kind: 'property' }] },
		] });
		expect(result?.constants.map(c => c.name)).toEqual(['Light', 'Dark']);
	});
	it('keeps qualified library enum types separate from local enum names', () => {
		const source = 'Enum VbMsgBoxResult\nCustom = 1\nEnd Enum';
		const call = `${source}\nSub T()\nDim answer As VBA.VbMsgBoxResult\nanswer = `;
		expect(resolveAssignmentValueCompletion(call, call.length)?.constants.map(c => c.name)).toContain('vbYes');
	});
	it('does not resolve an invalid module-qualified enum declaration', () => {
		const source = 'Enum Direction\nNorth = 1\nEnd Enum\nSub T()\nDim facing As Module.Direction\nfacing = ';
		expect(resolveAssignmentValueCompletion(source, source.length)).toBeUndefined();
	});
	it('handles leading-dot assignments inside With blocks', () => {
		const source = 'Sub T()\nWith ActiveCell\n.HorizontalAlignment = ';
		expect(resolveAssignmentValueCompletion(source, source.length)?.constants.map(c => c.name)).toContain('xlHAlignCenter');
	});
	it('does not complete a comparison passed to a procedure', () => {
		expect(accepted('Debug.Print ActiveCell.Interior.Color = ')).toBeUndefined();
	});
	it('supports escaped enum-valued properties', () => {
		expect(accepted('ActiveCell.[HorizontalAlignment] = ')?.constants.map(c => c.name)).toContain('xlHAlignCenter');
	});
	it('prioritizes values in single-line If and numbered assignments', () => {
		expect(accepted('If ok Then ActiveCell.HorizontalAlignment = ')?.constants.map(c => c.name)).toContain('xlHAlignCenter');
		expect(accepted('10 ActiveCell.HorizontalAlignment = ')?.constants.map(c => c.name)).toContain('xlHAlignCenter');
	});
	it('matches the named argument even with the caret inside an existing value', () => {
		const source = 'Sub T()\nThisWorkbook.BreakLink Type:=xlLinkTypeExcelLinks\nEnd Sub';
		const offset = source.indexOf('xlLinkType') + 3;
		expect(resolveArgumentValueCompletion(source, offset)?.enumName).toBe('XlLinkType');
	});
	it('prioritizes colors for the workbook cell chain', () => {
		expect(accepted('ThisWorkbook.Sheets(1).Cells(1).Interior.Color = ')?.constants.map(c => c.name)).toContain('vbRed');
	});
	it('offers enum members for a typed host property', () => {
		expect(accepted('ActiveCell.HorizontalAlignment = xl')?.constants.map(c => c.name)).toContain('xlHAlignCenter');
	});
	it('offers values for an enum-typed variable', () => {
		expect(accepted('alignment = ', 'Dim alignment As XlHAlign')?.constants.map(c => c.name)).toContain('xlHAlignCenter');
	});
	it('offers source-defined enum members', () => {
		expect(accepted('facing = ', 'Enum Direction\nNorth = 1\nSouth = 2\nEnd Enum\nDim facing As Direction')?.constants.map(c => c.name)).toEqual(['North', 'South']);
	});
	it('offers VBA runtime enums in assignments and arguments with defaults', () => {
		expect(accepted('answer = ', 'Dim answer As VbMsgBoxResult')?.constants.map(c => c.name)).toContain('vbYes');
		const source = 'Sub T()\nMsgBox "Continue?", ';
		expect(resolveArgumentValueCompletion(source, source.length)?.constants.map(c => c.name)).toContain('vbYesNo');
	});
	it('offers known enum values for Variant properties', () => {
		expect(accepted('ActiveCell.Interior.Pattern = ')?.constants.map(c => c.name)).toContain('xlPatternSolid');
		expect(accepted('ActiveCell.Borders(1).LineStyle = ')?.constants.map(c => c.name)).toContain('xlContinuous');
	});
	it.each(['If ActiveCell.HorizontalAlignment = ', 'x = "Color = ', "' ActiveCell.Color = ", 'x = y + ', 'x = 1: Debug.Print x = '])('ignores non-assignment slots: %s', line => {
		expect(accepted(line)).toBeUndefined();
	});
	it('opens suggestions after an assignment space', () => {
		expect(spaceTriggerMayComplete('ActiveCell.HorizontalAlignment = ')).toBe(true);
	});
});


describe('assignment trigger gates', () => {
 it.each(['.HorizontalAlignment =', 'If ok Then .HorizontalAlignment =', 'Let alignment =', 'Call Thing(Type:='])('accepts value slot %s', line => {
  expect(assignmentValueTriggerMayComplete(line)).toBe(true);
 });
 it.each(['If x =', 'Debug.Print x =', 'Set obj =', "' x =", 'x = "', 'x = 1 +', 'For i ='])('rejects unrelated equals %s', line => {
  expect(assignmentValueTriggerMayComplete(line)).toBe(false);
 });
});


afterEach(() => vi.restoreAllMocks());
describe('assignment target safety and work', () => {
 it('keeps host enums separate from a same-named local enum', () => {
  expect(accepted('ActiveCell.HorizontalAlignment = ', 'Enum XlHAlign\nBogus = 1\nEnd Enum')?.constants.map(c => c.name)).toContain('xlHAlignLeft');
 });
 it('offers Boolean values and handles logical continuations', () => {
  expect(accepted('enabled = ', 'Dim enabled As Boolean')?.constants.map(c => c.name)).toEqual(['True', 'False']);
  expect(accepted('ActiveCell.HorizontalAlignment = _\n ' )?.enumName).toBe('XlHAlign');
 });
 it('rejects read-only properties and methods', () => {
  for (const member of [{name:'State',kind:'property' as const,declaredType:'Boolean',writable:false}, {name:'State',kind:'method' as const,declaredType:'Boolean'}]) {
   const source = 'Sub T()\nDim item As Widget\nitem.State = ';
   expect(resolveAssignmentValueCompletion(source, source.length, {projectClassMembers:[{name:'Widget',moduleName:'Widget',kind:'class',members:[member]}]})).toBeUndefined();
  }
 });
 it('lexes only the assignment logical line even after thousands of statements', async () => {
  const lexer = await import('../src/analyzer/lexer/tokenize');
  const { assignmentTargetAt } = await import('../src/analyzer/completion/assignmentValueCompletion');
  const source = 'Sub WorkProbe()\n' + 'x = 1\n'.repeat(12000) + '.HorizontalAlignment = xlH';
  const tokenize = vi.spyOn(lexer, 'tokenize');
  expect(assignmentTargetAt(source, source.length)?.at(-1)?.rawText).toBe('HorizontalAlignment');
  expect(tokenize).toHaveBeenCalled();
  expect(Math.max(...tokenize.mock.calls.map(args => args[0].length))).toBeLessThan(100);
 });
});


describe('enum snapshot resolution', () => {
 it('resolves library-qualified host enums', () => {
  expect(accepted('chosen = ', 'Dim chosen As Excel.XlHAlign')?.enumName).toBe('XlHAlign');
  expect(accepted('chosen = ', 'Dim chosen As Missing.XlHAlign')).toBeUndefined();
 });
 it('avoids reparsing source enums when project surfaces are available', async () => {
  const builder=await import('../src/analyzer/symbols/buildModuleSymbols');
  const source='Sub EnumSnapshotProbe()\nDim facing As Direction\nfacing = ';
  const spy=vi.spyOn(builder,'buildModuleSymbols');
  const ctx={moduleName:'Module1',projectClassMembers:[{name:'Direction',moduleName:'Module1',kind:'enum' as const,members:[{name:'North',kind:'property' as const}]}]};
  expect(resolveAssignmentValueCompletion(source,source.length,ctx)?.constants.map(c=>c.name)).toEqual(['North']);
  // Expression binding may create its own initial snapshot. Repeat requests
  // should reuse it and never rebuild a separate graph for enum lookup.
  spy.mockClear();
  resolveAssignmentValueCompletion(source,source.length,ctx);
  expect(spy).not.toHaveBeenCalled();
 });
});

it('resolves a property enum in its declaring module, not a same-named caller enum', () => {
 const source='Sub T()\nDim item As Widget\nitem.State = ';
 const result=resolveAssignmentValueCompletion(source,source.length,{moduleName:'Caller',projectClassMembers:[
  {name:'Widget',moduleName:'Widget',kind:'class',members:[{name:'State',kind:'property',declaredType:'Status',writable:true}]},
  {name:'Status',moduleName:'Caller',kind:'enum',members:[{name:'CallerOnly',kind:'property'}]},
  {name:'Status',moduleName:'Widget',kind:'enum',members:[{name:'WidgetOnly',kind:'property'}]},
 ]});
 expect(result?.constants.map(c=>c.name)).toEqual(['WidgetOnly']);
});
