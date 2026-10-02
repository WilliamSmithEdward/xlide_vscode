// Diagnostics tests: Word and PowerPoint objects read as values (issue #462).
// Measured in Word and PowerPoint 16.0 64-bit (2026-10-02) through
// pyVBAharness: a never-set object read as an operand raises 91; Paragraphs,
// Tables and Slides need an index for their default member Item; a
// Paragraph's default member Range holds an object.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const SET: Readonly<Record<string, string>> = {
	Range: 'Set x = ActiveDocument.Range(0, 0)',
	Selection: 'Set x = Selection',
	Paragraphs: 'Set x = ActiveDocument.Paragraphs',
	Tables: 'Set x = ActiveDocument.Tables',
	Paragraph: 'Set x = ActiveDocument.Paragraphs(1)',
	TextRange: 'Set x = ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange',
	Slides: 'Set x = ActivePresentation.Slides',
};

function found(host: 'word' | 'powerpoint', type: string, set: boolean, use: string): string[] {
	const lines = [`Dim x As ${type}`, ...(set ? [SET[type]] : []), use];
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src, { host }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a never-set Word or PowerPoint object read as an operand', () => {
	it('raises 91', () => {
		for (const [host, type] of [['word', 'Range'], ['word', 'Selection'], ['powerpoint', 'TextRange']] as const) {
			for (const use of ['Main = x + 1', 'Main = x & "a"']) {
				expect(found(host, type, false, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^object-variable-not-set: .*'91'/)]);
			}
		}
	});

	it('runs once the object is set', () => {
		expect(found('word', 'Range', true, 'Main = x & "a"')).toEqual([]);
		expect(found('powerpoint', 'TextRange', true, 'Main = x & "a"')).toEqual([]);
	});
});

describe('Paragraphs, Tables and Slides read as a value', () => {
	it('raise 450 read whole', () => {
		for (const [host, type] of [['word', 'Paragraphs'], ['word', 'Tables'], ['powerpoint', 'Slides']] as const) {
			expect(found(host, type, true, 'Main = x'), type).toEqual([expect.stringMatching(/^object-default-value: .*'450'/)]);
		}
	});

	it('are Argument not optional as an operand, set or not', () => {
		for (const [host, type] of [['word', 'Paragraphs'], ['word', 'Tables'], ['powerpoint', 'Slides']] as const) {
			for (const use of ['Main = x + 1', 'Main = x & "a"']) {
				expect(found(host, type, true, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^collection-operand: .*Argument not optional/)]);
			}
		}
		expect(found('word', 'Paragraphs', false, 'Main = x + 1')).toEqual([expect.stringMatching(/^collection-operand: .*Argument not optional/)]);
	});
});

describe('a Paragraph, whose default member holds an object', () => {
	it('refuses a Let, an operator and a Let into a String while compiling', () => {
		expect(found('word', 'Paragraph', true, 'x = 5')).toEqual([expect.stringMatching(/^invalid-property-use: .*Invalid use of property/)]);
		expect(found('word', 'Paragraph', false, 'x = 5')).toEqual([expect.stringMatching(/^invalid-property-use: /)]);
		for (const use of ['Main = x + 1', 'Main = x & "a"', 'Dim s As String: s = x']) {
			expect(found('word', 'Paragraph', true, use), use).toEqual([expect.stringMatching(/^collection-operand: .*Type mismatch/)]);
		}
	});

	it('gives its Range read whole into a Variant', () => {
		expect(found('word', 'Paragraph', true, 'Main = x')).toEqual([]);
	});
});
