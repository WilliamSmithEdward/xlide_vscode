// Diagnostics tests: built-in arguments one step past the ones checked
// before - DateAdd below year 100, first-day-of-week and first-week-of-year
// bounds, FormatDateTime's NamedFormat, times out of range, file numbers past
// 512, Seek 0, Open's Len, empty paths, and reading a file created empty
// (issue #262). Every case was measured in Excel 16.0 (build 20326,
// 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim f As String, s As String\n    f = Environ$("TEMP") & "\\x.txt"\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('date arguments', () => {
	it.each([
		['Main = DateAdd("yyyy", -1, #1/1/100#)', '#1/1/100#', 'carry before January 1, 100'],
		['Main = DateAdd("d", -1, #1/1/100#)', '#1/1/100#', 'the -1 d interval(s)'],
		['Main = DateAdd("q", -1, #2/1/100#)', '#2/1/100#', 'before January 1, 100'],
		['Main = DateAdd("ww", -1, #1/7/100#)', '#1/7/100#', 'before January 1, 100'],
		['Main = DateAdd("h", -1, #1/1/100#)', '#1/1/100#', 'before January 1, 100'],
		['Main = DateAdd("s", -1, #1/1/100#)', '#1/1/100#', 'before January 1, 100'],
		['Main = DateAdd("d", -1, DateSerial(100, 1, 1))', 'DateSerial(100, 1, 1)', 'before January 1, 100'],
		['Main = DateAdd("h", 24, #12/31/9999#)', '#12/31/9999#', 'past December 31, 9999'],
	])('reports %s', (line, span, message) => {
		const src = source(line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { span, message });
	});

	it.each([
		['Main = DatePart("ww", #1/1/2000#, 8)', '8', "'FirstDayOfWeek' of 'DatePart' is 8"],
		['Main = DatePart("ww", #1/1/2000#, -1)', '-1', "'FirstDayOfWeek' of 'DatePart' is -1"],
		['Main = DatePart("ww", #1/1/2000#, 1, 4)', '4', "'FirstWeekOfYear' of 'DatePart' is 4"],
		['Main = DatePart("ww", #1/1/2000#, FirstDayOfWeek:=8)', '8', "'FirstDayOfWeek' of 'DatePart' is 8"],
		['Main = Format$(#1/1/2000#, "ww", 9)', '9', "'FirstDayOfWeek' of 'Format$' is 9"],
		['Main = Format(#1/1/2000#, "ww", 1, -1)', '-1', "'FirstWeekOfYear' of 'Format' is -1"],
		['Main = Format$(5, "0", 9)', '9', "'FirstDayOfWeek'"],
		['Main = FormatDateTime(#1/1/2000#, 9)', '9', "'NamedFormat' of 'FormatDateTime' is 9"],
		['Main = FormatDateTime(#1/1/2000#, -1)', '-1', "'NamedFormat' of 'FormatDateTime' is -1"],
		['Main = FormatDateTime(#1/1/2000#, 5)', '5', "'NamedFormat' of 'FormatDateTime' is 5"],
		['Main = DateDiff("ww", #1/1/2000#, #2/1/2000#, 8)', '8', "'FirstDayOfWeek' of 'DateDiff' is 8"],
	])('reports %s', (line, span, message) => {
		const src = source(line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { span, message });
	});

	it.each([
		['Main = TimeValue("25:00")', '"25:00"'],
		['Main = TimeValue("24:00")', '"24:00"'],
		['Main = TimeValue("10:60")', '"10:60"'],
		['Main = TimeValue("10:00:60")', '"10:00:60"'],
		['Main = TimeValue("1:2:3:4")', '"1:2:3:4"'],
		['Main = TimeValue("10")', '"10"'],
		['Main = CDate("25:00")', '"25:00"'],
		['Main = DateValue("25:00")', '"25:00"'],
	])('reports %s', (line, span) => {
		const src = source(line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-conversion-value'), 'runtime-conversion-value', { span, message: `cannot convert ${span} to Date` });
	});

	it('stays quiet in range', () => {
		for (const line of [
			'Main = DateAdd("yyyy", 1, #1/1/100#)',
			'Main = DateAdd("ww", -1, #1/8/100#)',
			'Main = DateAdd("m", -1, #3/31/100#)',
			'Main = DateAdd("yyyy", -1900, #1/1/2000#)',
			'Main = DateAdd("h", 1, #12/31/9999#)',
			'Main = DateAdd("h", -1, #1/1/100 10:00#)',
			'Main = DateAdd("d", -0.6, #1/1/100#)',
			'Main = DatePart("ww", #1/1/2000#, 0)',
			'Main = DatePart("ww", #1/1/2000#, 7, 3)',
			'Main = Format$(#1/1/2000#, "ww", 7, 3)',
			'Main = DateDiff("ww", #1/1/2000#, #2/1/2000#, 1, 4)',
			'Main = FormatDateTime(#1/1/2000#, 4)',
			'Main = FormatDateTime(#1/1/2000#, vbLongTime)',
			'Main = TimeValue("23:59:59")',
			'Main = TimeValue("13:00 PM")',
			'Main = TimeValue("10.5")',
			'Main = TimeValue(" 10:00 ")',
			'Main = CDate("10")',
		]) {
			expect(errors(source(line)), line).toEqual([]);
		}
	});
});

describe('file numbers, records and Len', () => {
	it.each([
		['Open f For Output As #513', '513', 'cannot be opened: file numbers run from 1 to 512'],
		['Open f For Output As 600', '600', 'cannot be opened'],
		['Open f For Output As #-1', '-1', 'cannot be opened'],
		['Close #600', '600', 'File number 600 is never open'],
		['Close #-1', '-1', 'File number -1 is never open'],
		['Main = LOF(600)', '600', 'File number 600 is never open'],
		['Print #1000, "x"', '1000', 'File number 1000 is never open'],
	])('reports %s', (line, span, message) => {
		const src = source(line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'file-number-zero'), 'file-number-zero', { span, message });
	});

	it('keeps the message for 0', () => {
		const src = source('Open f For Output As #0');
		expectDiagnostic(src, byCode(analyzeModule(src), 'file-number-zero'), 'file-number-zero', { span: '0', message: 'File number 0 cannot be opened: file numbers run from 1 to 512' });
	});

	it.each([
		[['Open f For Output As #7', 'Seek #7, 0', 'Close #7'], '0'],
		[['Open f For Output As #7', 'Seek #7, -1', 'Close #7'], '-1'],
		[['Seek #7, 0'], '0'],
		[['Open f For Binary As #7', 'Seek #7, 0', 'Close #7'], '0'],
	])('reports Seek below 1: %j', (lines, span) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'file-record-zero'), 'file-record-zero', { span, message: `Seek with record number ${span}` });
	});

	it.each([
		['Open f For Random As #7 Len = 0', 'runtime-argument-value', "Argument 'Len' of 'Open' is 0"],
		['Open f For Output As #7 Len = 0', 'runtime-argument-value', "Argument 'Len' of 'Open' is 0"],
		['Open f For Random As #7 Len = 32768', 'arithmetic-overflow', "Open's Len of 32768 does not fit an Integer"],
		['Open f For Random As #7 Len = 100000', 'arithmetic-overflow', "Open's Len of 100000"],
	])('reports %s', (line, code, message) => {
		const src = source(line, 'Close #7');
		expectDiagnostic(src, byCode(analyzeModule(src), code), code, { message });
	});

	it('stays quiet in range', () => {
		for (const lines of [
			['Open f For Output As #512', 'Print #512, "x"', 'Close #512'],
			['Open f For Output As #511', 'Close #511'],
			['Close #512'],
			['Open f For Output As #7', 'Seek #7, 1', 'Close #7'],
			['Open f For Random As #7 Len = 1', 'Close #7'],
			['Open f For Random As #7 Len = 32767', 'Close #7'],
			['Open f For Random As #7 Len = -1', 'Close #7'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});
});

describe('empty paths', () => {
	it.each([
		['Open "" For Input As #7', '""', "Open is given an empty path. This will raise Run-time error '75': Path/File access error."],
		['Open " " For Input As #7', '" "', "Open is given a path of spaces. This will raise Run-time error '53': File not found."],
		['Main = FileLen("")', '""', "FileLen is given an empty path. This will raise Run-time error '53'"],
		['Main = FileLen(" ")', '" "', "FileLen is given a path of spaces"],
		['Main = FileDateTime("")', '""', "'53'"],
		['Main = GetAttr("")', '""', "'53'"],
		['Main = Dir(" ")', '" "', "Dir is given a path of spaces. This will raise Run-time error '53'"],
		['MkDir ""', '""', "MkDir is given an empty path. This will raise Run-time error '76': Path not found."],
		['MkDir " "', '" "', "'76'"],
		['ChDir ""', '""', "'76'"],
		['ChDir " "', '" "', "'76'"],
		['RmDir ""', '""', "'76'"],
		['Kill ""', '""', "'53'"],
		['SetAttr "", vbNormal', '""', "'53'"],
		['FileCopy "", f', '""', "'75'"],
		['FileCopy f, ""', '""', "'75'"],
		['Name "" As f', '""', "'75'"],
		['Dim p As String\n    MkDir p', 'p', "MkDir is given 'p', which holds \"\" here"],
		['Const P As String = ""\n    Kill P', 'P', "Kill is given 'P', which holds \"\" here"],
	])('reports %s', (line, span, message) => {
		const src = source(line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'empty-file-path'), 'empty-file-path', { span, message });
	});

	it('stays quiet on a path, Dir(""), and the cases not measured', () => {
		for (const line of [
			'Main = Len(Dir("")) >= 0',
			'ChDrive ""',
			'Main = FileLen(f)',
			'Kill " "',
			'Name f As ""',
			'Dim p As String\n    p = "c:\\x"\n    MkDir p',
		]) {
			expect(errors(source(line)), line).not.toContain('empty-file-path');
		}
	});
});

describe('reading a file created empty', () => {
	const created = ['Open f For Output As #7', 'Close #7'];

	it.each([
		[[...created, 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'], 'Line', "'Line Input #' reads file #7"],
		[[...created, 'Open f For Input As #7', 'Input #7, s', 'Close #7'], 'Input', "'Input #' reads file #7"],
		[[...created, 'Open f For Input As #7', 's = Input(1, #7)', 'Close #7'], 'Input', "'Input' reads file #7"],
		[[...created, 'Open f For Input As #8', 'Line Input #8, s', 'Close #8'], 'Line', 'reads file #8'],
		[['Open f For Output As #7', 'Close', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'], 'Line', 'reads file #7'],
		[['Open f For Output As #7', 'Print #7, "";', 'Close #7', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'], 'Line', 'reads file #7'],
	])('reports %j', (lines, span, message) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'file-read-past-end'), 'file-read-past-end', { span, message: `${message}, which this procedure created empty and reopened For Input without checking EOF. This will raise Run-time error '62'` });
	});

	it('stays quiet once something is written, checked, or the path changes', () => {
		for (const lines of [
			['Open f For Output As #7', 'Print #7, "a"', 'Close #7', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			['Open f For Output As #7', 'Print #7,', 'Close #7', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			['Open f For Output As #7', 'Write #7,', 'Close #7', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			[...created, 'Open f For Input As #7', 'If Not EOF(7) Then Line Input #7, s', 'Close #7'],
			[...created, 'Open f For Input As #7', 'Do Until EOF(7)', '    Line Input #7, s', 'Loop', 'Close #7'],
			[...created, 'Open f For Input As #7', 'Main = LOF(7)', 'Line Input #7, s', 'Close #7'],
			[...created, 'Open f For Append As #7', 'Print #7, "a"', 'Close #7', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			[...created, 'f = f & ".b"', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			[...created, 'Fill f', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			[...created, 's = Twist(f)', 'Open f For Input As #7', 'Line Input #7, s', 'Close #7'],
			[...created, 'Open f For Binary As #7', 'Get #7, , s', 'Close #7'],
		]) {
			const src = `${source(...lines)}Private Sub Fill(ByRef p As String)\n    p = p & ".b"\nEnd Sub\nPrivate Function Twist(ByRef p As String) As String\n    p = p & ".b"\nEnd Function\n`;
			expect(errors(src), lines.join(': ')).not.toContain('file-read-past-end');
		}
	});
});
