// Diagnostics tests: late-bound objects whose faults the code's literals make
// plain (issue #477). Each case was run through pyVBAharness on 2026-10-02 in
// Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { parseXml, vbscriptPatternError } from '../../src/analyzer/diagnostics/rules/lateBoundObjects';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const RE = 'Dim re As Object\n    Set re = CreateObject("VBScript.RegExp")\n    ';
const FSO = 'Dim fso As Object, ts As Object, p As String\n    Set fso = CreateObject("Scripting.FileSystemObject")\n    p = Environ("TEMP") & "\\x.txt"\n    If fso.FileExists(p) Then fso.DeleteFile p\n    ';
const RS = 'Dim rs As Object\n    Set rs = CreateObject("ADODB.Recordset")\n    ';
const DOC = 'Dim doc As Object\n    Set doc = CreateObject("MSXML2.DOMDocument.6.0")\n    ';

describe('ProgIDs (issue #477)', () => {
	it('reports one that names no class', () => {
		for (const id of ['Scripting.Dictionery', 'Dictionary', '', 'RegExp', 'VBScript.RegEx', 'Scripting.FileSystemObjects']) {
			expect(errors(`Dim o As Object\n    Set o = CreateObject("${id}")`), id).toEqual(['runtime-argument-value']);
		}
		expect(errors('Dim o As Object\n    Set o = GetObject(, "Excel.Aplication")')).toEqual(['runtime-argument-value']);
	});

	it('takes any case, a version and a ProgID it does not know', () => {
		for (const id of ['scripting.dictionary', 'MSXML2.DOMDocument.6.0', 'MSXML2.DOMDocument.3.0', 'htmlfile']) {
			expect(errors(`Dim o As Object\n    Set o = CreateObject("${id}")`), id).toEqual([]);
		}
	});
});

describe('VBScript patterns (issue #477)', () => {
	it('gives the error VBScript raises', () => {
		const cases: Array<[string, number | undefined]> = [
			['(a', 5020], ['[a', 5019], ['*a', 5018], ['a++', 5018], ['a+??', 5018], ['^*', 5018], ['\\b*', 5018], ['x|*', 5018],
			['(?<=a)b', 5017], ['a)', 5017], ['a{2,1}', 5017], ['a\\', 5017], ['(?<n>a)', 5017], ['[^]', 5019],
			['a??', undefined], ['a+?', undefined], ['a{', undefined], ['(?!a)b', undefined], ['[]a]', undefined], ['(?:a|b)+\\d{2}', undefined],
		];
		for (const [pattern, error] of cases) {
			expect(vbscriptPatternError(pattern), pattern).toBe(error);
		}
	});

	it('reports a bad pattern where it is used, not where it is set', () => {
		expect(errors(`${RE}re.Pattern = "(a"\n    Main = re.Test("a")`)).toEqual(['runtime-argument-value']);
		expect(errors(`${RE}re.Pattern = "(a"\n    Main = 1`)).toEqual([]);
	});
});

describe('a RegExp used (issue #477)', () => {
	it('reports a member it lacks, Null, and text for a Boolean', () => {
		expect(errors(`${RE}re.Patern = "a"`)).toEqual(['runtime-member-not-found']);
		expect(errors(`${RE}Main = re.Test(Null)`)).toEqual(['runtime-argument-value']);
		expect(errors(`${RE}re.Global = "abc"`)).toEqual(['runtime-argument-value']);
		expect(errors(`${RE}re.Global = "True"\n    Main = re.Global`)).toEqual([]);
	});

	it('reports a match or group past the last', () => {
		expect(errors(`${RE}re.Pattern = "z"\n    Main = re.Execute("abc")(0).Value`)).toEqual(['collection-index-out-of-range']);
		expect(errors(`${RE}re.Pattern = "(a)"\n    Main = re.Execute("a")(0).SubMatches(1)`)).toEqual(['collection-index-out-of-range']);
		expect(errors(`${RE}re.Pattern = "a"\n    Main = re.Execute("aXa")(1).Value`)).toEqual(['collection-index-out-of-range']);
		expect(errors(`${RE}re.Global = True\n    re.Pattern = "a"\n    Main = re.Execute("aXa")(1).Value`)).toEqual([]);
		expect(errors(`${RE}re.IgnoreCase = True\n    re.Pattern = "a"\n    Main = re.Execute("A")(0).Value`)).toEqual([]);
		expect(errors(`${RE}re.Pattern = "z"\n    Main = re.Execute("abc").Count`)).toEqual([]);
	});
});

describe('a FileSystemObject and its TextStreams (issue #477)', () => {
	it('reports a mode the stream refuses, and a stream used after Close', () => {
		expect(errors(`${FSO}Set ts = fso.OpenTextFile(p, 3, True)`)).toEqual(['runtime-argument-value']);
		expect(errors(`${FSO}Set ts = fso.CreateTextFile(p, True)\n    Main = ts.ReadLine`)).toEqual(['file-mode-mismatch']);
		expect(errors(`${FSO}Set ts = fso.CreateTextFile(p, True)\n    ts.Close\n    ts.WriteLine "y"`)).toEqual(['object-variable-not-set']);
		expect(errors(`${FSO}Set ts = fso.CreateTextFile(p, True)\n    ts.Close\n    ts.Close`)).toEqual([]);
	});

	it('follows what the procedure did to the path', () => {
		expect(errors(`${FSO}fso.CreateTextFile(p, True).Close\n    Set ts = fso.OpenTextFile(p, 1)\n    Main = ts.ReadLine`)).toEqual(['file-mode-mismatch']);
		expect(errors(`${FSO}fso.CreateTextFile(p, True).Close\n    fso.CreateTextFile p, False`)).toEqual(['runtime-argument-value']);
		expect(errors(`${FSO}Main = fso.GetFile(p).Size`)).toEqual(['runtime-argument-value']);
		expect(errors(`${FSO}fso.DeleteFile p`)).toEqual(['runtime-argument-value']);
		expect(errors(`${FSO}Set ts = fso.CreateTextFile(p, True)\n    ts.WriteLine "x"\n    ts.Close\n    Set ts = fso.OpenTextFile(p, 1)\n    Main = ts.ReadLine`)).toEqual([]);
		expect(errors(`${FSO}fso.CreateTextFile(p, True).Close\n    fso.CreateTextFile p`)).toEqual([]);
		expect(errors(`${FSO}p = "C:\\other.txt"\n    Main = fso.GetFile(p).Size`)).toEqual([]);
		expect(errors(`${FSO}If fso.FileExists(p) Then fso.DeleteFile p`)).toEqual([]);
	});
});

describe('a Recordset never opened (issue #477)', () => {
	it('refuses its records', () => {
		for (const use of ['rs.MoveNext', 'Main = rs.EOF', 'rs.Close', 'Main = rs.RecordCount', 'rs.MoveLast']) {
			expect(errors(`${RS}${use}`), use).toEqual(['late-bound-object-state']);
		}
		expect(errors(`${RS}Main = rs.State`)).toEqual([]);
		expect(errors(`${RS}Main = rs.Fields.Count`)).toEqual([]);
	});
});

describe('an MSXML2.DOMDocument (issue #477)', () => {
	it('reports a document with no element, and a path no element is on', () => {
		expect(errors(`${DOC}Main = doc.DocumentElement.NodeName`)).toEqual(['object-variable-not-set']);
		expect(errors(`${DOC}doc.LoadXML "<a>"\n    Main = doc.DocumentElement.NodeName`)).toEqual(['object-variable-not-set']);
		expect(errors(`${DOC}doc.LoadXML "<a/>"\n    Main = doc.SelectSingleNode("//b").Text`)).toEqual(['object-variable-not-set']);
		expect(errors(`${DOC}doc.LoadXML "<A/>"\n    Main = doc.SelectSingleNode("//a").Text`)).toEqual(['object-variable-not-set']);
		expect(errors(`${DOC}doc.LoadXML "<a/>"\n    Main = doc.SelectNodes("//[").Length`)).toEqual(['runtime-argument-value']);
	});

	it('stays quiet where the element is there', () => {
		expect(errors(`${DOC}doc.LoadXML "<a><b>t</b></a>"\n    Main = doc.SelectSingleNode("//b").Text`)).toEqual([]);
		expect(errors(`${DOC}doc.LoadXML "<a/>"\n    Main = doc.SelectSingleNode("a").NodeName`)).toEqual([]);
		expect(errors(`${DOC}doc.LoadXML "<a/>"\n    Main = doc.SelectSingleNode("//b") Is Nothing`)).toEqual([]);
	});

	it('reads XML as strictly as LoadXML', () => {
		for (const bad of ['<a>', '<a></b>', '<a/><b/>', '<a x=1/>', '<a>&</a>']) {
			expect(parseXml(bad), bad).toBeNull();
		}
		for (const good of ['<a/>', '<a x="1"><b>t</b></a>', '<?xml version="1.0"?><a><!-- c --><![CDATA[<x>]]></a>', '<a>&amp;</a>']) {
			expect(parseXml(good), good).not.toBeNull();
		}
	});
});
