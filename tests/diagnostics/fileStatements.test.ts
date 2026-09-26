// Diagnostics tests: file statements whose failure the code proves (issue
// #123). Each raising sample was measured in Excel 16.0 (build 20326,
// 2026-09-26); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
}

describe('file-number-zero (issue #123)', () => {
	it('flags As #0 and LOF(0)', () => {
		const src = wrap('Open Environ$("TEMP") & "\\a.txt" For Output As #0', 'Main = LOF(0)');
		expectDiagnostics(src, analyzeModule(src), 'file-number-zero', [
			{ span: '0', message: "error '52'" },
			{ span: '0' },
		]);
	});
});

describe('file-used-after-close (issue #123)', () => {
	it('flags Print # on a number closed above', () => {
		const src = wrap(
			'Dim f As Integer',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'Close #f',
			'Print #f, "late"',
		);
		expectDiagnostic(src, analyzeModule(src), 'file-used-after-close', { span: 'f', message: "error '52'" });
	});

	it('stays quiet when the number is opened again, or a block stands between', () => {
		const reopened = wrap(
			'Dim f As Integer',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'Close #f',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'Print #f, "again"',
			'Close #f',
		);
		const blocked = wrap(
			'Dim f As Integer',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'Close #f',
			'If Main = 0 Then',
			'    Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'End If',
			'Print #f, "maybe"',
		);
		expect(byCode(analyzeModule(reopened), 'file-used-after-close')).toHaveLength(0);
		expect(byCode(analyzeModule(blocked), 'file-used-after-close')).toHaveLength(0);
	});
});

describe('file-mode-mismatch (issue #123)', () => {
	it('flags Print # on an Input file and Line Input # on an Output file', () => {
		const src = wrap(
			'Dim f As Integer, s As String',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.txt" For Input As #f',
			'Print #f, "x"',
			'Close #f',
			'Open Environ$("TEMP") & "\\b.txt" For Output As #7',
			'Line Input #7, s',
			'Close #7',
		);
		expectDiagnostics(src, analyzeModule(src), 'file-mode-mismatch', [
			{ span: 'Print', message: ['For Input', "error '54'"] },
			{ span: 'Line', message: 'For Output' },
		]);
	});

	it('stays quiet for the matching mode and for Random or Binary files', () => {
		const src = wrap(
			'Dim f As Integer, s As String, x As Long',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'Print #f, "x"',
			'Close #f',
			'Open Environ$("TEMP") & "\\a.txt" For Input As #f',
			'Line Input #f, s',
			'Close #f',
			'Open Environ$("TEMP") & "\\a.bin" For Binary As #f',
			'Put #f, 1, x',
			'Get #f, 1, x',
			'Close #f',
		);
		expect(byCode(analyzeModule(src), 'file-mode-mismatch')).toHaveLength(0);
	});
});

describe('file-already-open (issue #123)', () => {
	it('flags a second Open As the same literal number', () => {
		const src = wrap(
			'Open Environ$("TEMP") & "\\a.txt" For Output As #7',
			'Open Environ$("TEMP") & "\\b.txt" For Output As #7',
			'Close #7',
		);
		expectDiagnostic(src, analyzeModule(src), 'file-already-open', { span: '7', message: "error '55'" });
	});

	it('stays quiet after a Close, or when FreeFile is asked again', () => {
		const src = wrap(
			'Dim f As Integer',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #7',
			'Close #7',
			'Open Environ$("TEMP") & "\\b.txt" For Output As #7',
			'Close',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.txt" For Output As #f',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\b.txt" For Output As #f',
			'Close',
		);
		expect(byCode(analyzeModule(src), 'file-already-open')).toHaveLength(0);
	});
});

describe('file-record-zero (issue #123)', () => {
	it('flags Seek #f, 0 and Put #f, 0, x on a Binary file', () => {
		const src = wrap(
			'Dim f As Integer, x As Long',
			'f = FreeFile',
			'Open Environ$("TEMP") & "\\a.bin" For Binary As #f',
			'Seek #f, 0',
			'Put #f, 0, x',
			'Put #f, 1, x',
			'Close #f',
		);
		expectDiagnostics(src, analyzeModule(src), 'file-record-zero', [
			{ span: '0', message: ['Seek', "error '63'"] },
			{ span: '0', message: 'Put' },
		]);
	});
});
