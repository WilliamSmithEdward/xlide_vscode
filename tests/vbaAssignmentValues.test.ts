import { describe, expect, it } from 'vitest';
import { resolveAssignmentValueCompletion } from '../src/analyzer/completion/assignmentValueCompletion';
import { spaceTriggerMayComplete } from '../src/analyzer/completion/cursorContext';
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
	it('resolves a qualified enum in the current module', () => {
		const source = 'Enum Direction\nNorth = 1\nEnd Enum\nSub T()\nDim facing As Module.Direction\nfacing = ';
		expect(resolveAssignmentValueCompletion(source, source.length)?.constants.map(c => c.name)).toEqual(['North']);
	});
	it('handles leading-dot assignments inside With blocks', () => {
		const source = 'Sub T()\nWith ActiveCell\n.HorizontalAlignment = ';
		expect(resolveAssignmentValueCompletion(source, source.length)?.constants.map(c => c.name)).toContain('xlHAlignCenter');
	});
	it('does not prioritize colors for a comparison passed to a procedure', () => {
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
