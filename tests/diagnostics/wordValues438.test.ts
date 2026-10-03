// Diagnostics tests: Word and PowerPoint objects read as values, by their
// type libraries' default members (issue #438). Each case was run through
// pyVBAharness on 2026-10-02 in Word and PowerPoint 16.0 64-bit.

import { describe, it, expect } from 'vitest';
import { hostTokenForFileName } from '../../src/analyzer/host/hostRegistry';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

function errors(file: string, body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', host: hostTokenForFileName(file), referencedHosts: [] }).diagnostics
		.filter((diag) => diag.severity === 'error').map((diag) => String(diag.code));
}

const word = (body: string): string[] => errors('Doc1.docm', body);
const DOC = 'Dim x As Document\n    Set x = ActiveDocument\n    ';

describe('a Word Document (issue #438)', () => {
	it('gives its Name, which no Let reaches and is no number', () => {
		expect(word(`${DOC}x = 5`)).toEqual(['readonly-member-assignment']);
		expect(word('Dim x As Document\n    x = 5')).toEqual(['readonly-member-assignment']);
		expect(word(`${DOC}Main = x + 1`)).toEqual(['assignment-type-mismatch']);
		expect(word(`${DOC}If x Then Main = 2`)).toEqual(['assignment-type-mismatch']);
		for (const body of ['Main = x', 'Main = x & "a"', 'If x = "a" Then Main = 2', 'Main = CStr(x)']) {
			expect(word(`${DOC}${body}`)).toEqual([]);
		}
	});

	it('is read late through an Object', () => {
		const obj = 'Dim x As Object\n    Set x = ActiveDocument\n    ';
		expect(word(`${obj}x = 5`)).toEqual(['object-default-value']);
		expect(word(`${obj}Main = x(1)`)).toEqual(['object-default-value']);
		expect(word(`${obj}Main = x + 1`)).toEqual(['assignment-type-mismatch']);
		expect(word(`${obj}Main = x & "a"`)).toEqual([]);
	});
});

describe('an index on a default member that takes none (issue #438)', () => {
	it('does not compile, set or not', () => {
		for (const setup of ['Dim x As Document\n    ', 'Dim x As Range\n    ', 'Dim x As Selection\n    ', `${DOC}`, 'Dim x As Paragraph\n    Set x = ActiveDocument.Paragraphs(1)\n    ']) {
			expect(word(`${setup}Main = x(1)`)).toEqual(['argument-count']);
		}
		expect(word('Dim x(1) As Range\n    Set x(0) = ActiveDocument.Range(0, 0)\n    Main = x(0).Start')).toEqual([]);
	});
});

describe('a Paragraph and the Item collections (issue #438)', () => {
	it('gives CStr and Len no value', () => {
		const para = 'Dim x As Paragraph\n    Set x = ActiveDocument.Paragraphs(1)\n    ';
		expect(word(`${para}Main = CStr(x)`)).toEqual(['collection-operand']);
		expect(word(`${para}Main = Len(x)`)).toEqual(['collection-operand']);
		for (const setup of ['Dim x As Paragraphs\n    Set x = ActiveDocument.Paragraphs\n    ', 'Dim x As Tables\n    Set x = ActiveDocument.Tables\n    ']) {
			expect(word(`${setup}Main = CStr(x)`)).toEqual(['collection-operand']);
			expect(word(`${setup}If x Then Main = 2`)).toEqual(['object-default-value']);
		}
	});
});

describe('PowerPoint (issue #438)', () => {
	it('reads Slides as the Item collection it is', () => {
		const slides = 'Dim x As Slides\n    Set x = ActivePresentation.Slides\n    ';
		expect(errors('Pres.pptm', `${slides}If x Then Main = 2`)).toEqual(['object-default-value']);
		expect(errors('Pres.pptm', `${slides}Main = CStr(x)`)).toEqual(['collection-operand']);
	});
});
