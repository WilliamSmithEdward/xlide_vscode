// Diagnostics tests: leftovers collected in issue #685 from closed issues.
// Each was measured on 2026-10-03 in Excel 16.0 (build 20430) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

function errors(body: string, after = ''): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n${after}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a fraction into a whole-number local (#673)', () => {
	it('is rounded half to even where a subscript reads it', () => {
		expect(errors('Dim a As Long, r(3) As Long\n    a = 4.4\n    Main = r(a)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim a As Long, r(3) As Long\n    a = 3.4\n    Main = r(a)')).toEqual([]);
		expect(errors('Dim a As Long, r(3) As Long\n    a = 3.5\n    Main = r(a)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim a As Double, r(3) As Long\n    a = 3.4\n    Main = r(a)')).toEqual([]);
	});
});

describe('Null in a logical operator (#407)', () => {
	it('still converts the other side to a Long', () => {
		expect(errors('Main = Null And 1E10')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = Null Or 1E10')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = 1E10 And Null')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = Null And 5')).toEqual([]);
		expect(errors('Main = Null Or Null')).toEqual([]);
	});
});

describe('Split with an empty delimiter, and a Static local (#559)', () => {
	it('gives the whole text as the one element', () => {
		expect(errors('Dim s As String\n    s = "a,b"\n    Main = Split(s, "")(1)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim s As String\n    s = "a,b"\n    Main = Split(s, "")(0)')).toEqual([]);
		expect(errors('Main = Split("", "")(0)')).toEqual([]);
	});

	it('reads a Static local the straight line assigned', () => {
		expect(errors('Static t As String\n    t = "a"\n    Main = Split(t, ",")(1)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Static t As String\n    Main = Split(t, ",")(1)\n    t = "a"')).toEqual([]);
	});
});

describe('a parameterless Function indexed (#609)', () => {
	const ARR = 'Private Function Arr() As Variant\n    Arr = Array(10, 20)\nEnd Function\n';
	it('reads the subscript against what it returns', () => {
		expect(errors('Main = Arr(5)', ARR)).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Main = Arr(1, 2)', ARR)).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Main = Arr(1)', ARR)).toEqual([]);
		expect(errors('Dim Arr(9) As Long\n    Main = Arr(5)', ARR)).toEqual([]);
	});
});

describe('For bounds read once (#346)', () => {
	const ARR = 'Dim arr(0 To 3) As Long, i As Long, n As Long\n    ';
	it('reads a bound the body writes as the loop starts', () => {
		expect(errors(`${ARR}n = 4\n    For i = 0 To n\n        n = 3\n        arr(i) = 1\n    Next\n    Main = 1`)).toEqual(['array-subscript-out-of-bounds']);
		expect(errors(`${ARR}n = 3\n    For i = 0 To n\n        n = 9\n        arr(i) = 1\n    Next\n    Main = 1`)).toEqual([]);
	});

	it('works out the last pass of a stepped symbolic bound', () => {
		expect(errors(`${ARR}n = 5\n    For i = 0 To n Step 2\n        arr(i) = 1\n    Next\n    Main = 1`)).toEqual(['array-subscript-out-of-bounds']);
		expect(errors(`${ARR}n = 3\n    For i = 0 To n Step 2\n        arr(i) = 1\n    Next\n    Main = 1`)).toEqual([]);
	});
});

describe('a callee that passes its parameter only to VBA functions (#449)', () => {
	it('leaves the caller\'s object as it was', () => {
		const touch = 'Private Sub TouchC(ByRef p As Collection)\n    Debug.Print TypeName(p)\nEnd Sub\n';
		expect(errors('Dim c As Collection\n    TouchC c\n    Main = c.Count', touch)).toEqual(['object-variable-not-set']);
		const sets = 'Private Sub TouchC(ByRef p As Collection)\n    Set p = New Collection\nEnd Sub\n';
		expect(errors('Dim c As Collection\n    TouchC c\n    Main = c.Count', sets)).toEqual([]);
		const passes = 'Private Sub TouchC(ByRef p As Collection)\n    Fill p\nEnd Sub\nPrivate Sub Fill(ByRef q As Collection)\n    Set q = New Collection\nEnd Sub\n';
		expect(errors('Dim c As Collection\n    TouchC c\n    Main = c.Count', passes)).toEqual([]);
	});
});

describe('late-bound objects (#477)', () => {
	const RE = 'Dim re As Object\n    Set re = CreateObject("VBScript.RegExp")\n    re.Global = True\n    re.Pattern = "^x"\n    ';
	const FS = 'Dim fso As Object, ts As Object\n    Set fso = CreateObject("Scripting.FileSystemObject")\n    Set ts = fso.CreateTextFile("C:\\t\\a.txt", True)\n    ';
	const RS = 'Dim rs As Object\n    Set rs = CreateObject("ADODB.Recordset")\n    rs.Fields.Append "a", 3\n    ';
	const DOC = 'Dim d As Object\n    Set d = CreateObject("MSXML2.DOMDocument.6.0")\n    ';
	it.each([
		['Execute past its matches with MultiLine off', `${RE}re.MultiLine = False\n    Main = re.Execute("x" & vbLf & "x")(1).Value`, 'collection-index-out-of-range'],
		['ReadAll of a file created empty', `${FS}ts.Close\n    Set ts = fso.OpenTextFile("C:\\t\\a.txt")\n    Main = ts.ReadAll`, 'file-mode-mismatch'],
		['a second Close', `${RS}rs.Open\n    rs.Close\n    rs.Close\n    Main = 1`, 'late-bound-object-state'],
		['EOF after Close', `${RS}rs.Open\n    rs.Close\n    Main = rs.EOF`, 'late-bound-object-state'],
		['an attribute no element has', `${DOC}d.LoadXML "<a><b id='1'/></a>"\n    Main = d.SelectSingleNode("//b[@id='2']").nodeName`, 'object-variable-not-set'],
		['a position past the siblings', `${DOC}d.LoadXML "<a><b/><b/></a>"\n    Main = d.SelectSingleNode("//b[3]").nodeName`, 'object-variable-not-set'],
		['a child step past them', `${DOC}d.LoadXML "<a><b/><b/></a>"\n    Main = d.SelectSingleNode("/a/b[3]").nodeName`, 'object-variable-not-set'],
	])('reports %s', (_label, body, code) => {
		expect(errors(body), body).toEqual([code]);
	});

	it.each([
		['Execute with MultiLine on', `${RE}re.MultiLine = True\n    Main = re.Execute("x" & vbLf & "x")(1).Value`],
		['ReadAll of a file written', `${FS}ts.WriteLine "a"\n    ts.Close\n    Set ts = fso.OpenTextFile("C:\\t\\a.txt")\n    Main = ts.ReadAll`],
		['one Close', `${RS}rs.Open\n    rs.Close\n    Main = 1`],
		['an attribute an element has', `${DOC}d.LoadXML "<a><b id='1'/></a>"\n    Main = d.SelectSingleNode("//b[@id='1']").nodeName`],
		['a position among its siblings', `${DOC}d.LoadXML "<a><b/><b/></a>"\n    Main = d.SelectSingleNode("//b[2]").nodeName`],
		['a second b under one parent', `${DOC}d.LoadXML "<a><c><b/></c><c><b/><b/></c></a>"\n    Main = d.SelectSingleNode("//b[2]").nodeName`],
	])('stays quiet on %s', (_label, body) => {
		expect(errors(body), body).toEqual([]);
	});
});

describe('a ParamArray passed on (#445)', () => {
	const K1 = 'Option Explicit\nPublic Function TakeRef(ByRef v As Variant) As Long\n    TakeRef = 1\nEnd Function\nPublic Function TakeVal(ByVal v As Variant) As Long\n    TakeVal = 1\nEnd Function\n';
	function found(body: string): string[] {
		const main = 'Option Explicit\nFunction Main() As Variant\n    Main = F(1, 2)\nEnd Function\n'
			+ `Private Function F(ParamArray p() As Variant) As Long\n    ${body}\nEnd Function\n`
			+ 'Private Function ByRefV(ByRef v As Variant) As Long\n    ByRefV = 1\nEnd Function\n'
			+ 'Private Function ByValV(ByVal v As Variant) As Long\n    ByValV = 1\nEnd Function\n';
		const project = buildVbaProjectIndex([
			{ moduleName: 'Module1', type: 'standard', source: main },
			{ moduleName: 'Module2', type: 'standard', source: K1 },
			{ moduleName: 'K1', type: 'class', source: K1 },
		]);
		return analyzeVbaModuleSource({ source: main, moduleName: 'Module1', moduleKind: 'standard', ...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)) } as never)
			.diagnostics.filter((diag) => diag.code === 'invalid-paramarray-use').map((diag) => diag.message);
	}

	it.each([
		['to another module\'s ByRef parameter', 'F = Module2.TakeRef(p)'],
		['to a class\'s ByRef parameter', 'Dim k As New K1: F = k.TakeRef(p)'],
		['by name to a ByRef parameter', 'F = ByRefV(v:=p)'],
		['to a late-bound member', 'Dim k As Object: Set k = New K1: F = k.TakeRef(p)'],
	])('reports it passed %s', (_label, body) => {
		expect(found(body), body).toHaveLength(1);
	});

	it.each([
		'F = Module2.TakeVal(p)',
		'Dim k As New K1: F = k.TakeVal(p)',
		'F = ByValV(v:=p)',
		'Dim k As Object: Set k = New K1: F = k.TakeRef(p(0))',
	])('stays quiet on %s', (body) => {
		expect(found(body), body).toEqual([]);
	});
});

describe('a late-bound class instance (#414)', () => {
	const K1 = 'Option Explicit\nPublic Property Get Idx(ByVal i As Long) As Long\n    Idx = i\nEnd Property\n'
		+ 'Public Property Get Opt(Optional ByVal i As Long = 2) As Long\n    Opt = i\nEnd Property\n'
		+ 'Public Function Need(ByVal i As Long) As Long\n    Need = i\nEnd Function\n'
		+ 'Public Property Get O() As Collection\n    Set O = New Collection\nEnd Property\n'
		+ 'Public Property Set O(ByVal c As Collection)\nEnd Property\n';
	function found(body: string): string[] {
		const project = buildVbaProjectIndex([
			{ moduleName: 'Module1', type: 'standard', source: `Option Explicit\nFunction Main() As Variant\n    Dim o As Object\n    Set o = New K1\n    ${body}\nEnd Function\n` },
			{ moduleName: 'K1', type: 'class', source: K1 },
		]);
		const src = `Option Explicit\nFunction Main() As Variant\n    Dim o As Object\n    Set o = New K1\n    ${body}\nEnd Function\n`;
		return analyzeVbaModuleSource({ source: src, moduleName: 'Module1', moduleKind: 'standard', ...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)) } as never)
			.diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
	}

	it('reports a member read with a required argument missing, and a Let with no Let', () => {
		expect(found('Main = o.Idx')).toEqual(['runtime-member-not-found']);
		expect(found('Main = o.Need')).toEqual(['runtime-member-not-found']);
		expect(found('o.O = New Collection\n    Main = 1')).toEqual(['runtime-member-not-found']);
	});

	it('stays quiet with the argument, an Optional one, and Set', () => {
		expect(found('Main = o.Idx(3)')).toEqual([]);
		expect(found('Main = o.Opt')).toEqual([]);
		expect(found('Main = o.Need(4)')).toEqual([]);
		expect(found('Set o.O = New Collection\n    Main = 1')).toEqual([]);
	});
});

describe('a Dictionary where a Collection goes (#410)', () => {
	const D = 'Dim o As Object\n    Set o = CreateObject("Scripting.Dictionary")\n    ';
	it('is refused by a Collection parameter and a Collection variable', () => {
		expect(errors(`${D}Main = TakeC(o)`, 'Private Function TakeC(c As Collection) As Long\n    TakeC = 1\nEnd Function\n')).toEqual(['argument-type-mismatch']);
		expect(errors(`${D}Dim c As Collection\n    Set c = o\n    Main = 1`)).toEqual(['assignment-object-type-mismatch']);
	});

	it('goes into Object and Variant', () => {
		expect(errors(`${D}Main = TakeO(o)`, 'Private Function TakeO(c As Object) As Long\n    TakeO = 1\nEnd Function\n')).toEqual([]);
		expect(errors(`${D}Main = TakeV(o)`, 'Private Function TakeV(c As Variant) As Long\n    TakeV = 1\nEnd Function\n')).toEqual([]);
	});
});

describe('a Rows.Count an Integer cannot hold (#411)', () => {
	const TAKEIR = 'Private Function TakeIR(i As Integer) As Long\n    TakeIR = i\nEnd Function\n';
	const TAKEI = 'Private Function TakeI(ByVal i As Integer) As Long\n    TakeI = i\nEnd Function\n';
	it('reads a With on a range', () => {
		expect(errors('Dim i As Integer\n    With ActiveSheet.Range("A1:A40000")\n        i = .Rows.Count\n    End With\n    Main = i')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim i As Integer\n    With ActiveSheet.Range("A1:A400")\n        i = .Rows.Count\n    End With\n    Main = i')).toEqual([]);
	});

	it('passes an expression to a ByRef Integer as a temporary Integer', () => {
		expect(errors('Main = TakeIR(Rows.Count)', TAKEIR)).toEqual(['arithmetic-overflow']);
		expect(errors('Main = TakeIR(Rows.Count \\ 100)', TAKEIR)).toEqual([]);
	});

	it('reads a With member passed as an argument', () => {
		expect(errors('With ActiveSheet\n        Main = TakeI(.Rows.Count)\n    End With', TAKEI)).toEqual(['arithmetic-overflow']);
		expect(errors('With ActiveSheet.Range("A1:A40000")\n        Main = TakeI(.Rows.Count)\n    End With', TAKEI)).toEqual(['arithmetic-overflow']);
		expect(errors('With ActiveSheet.Range("A1:C3")\n        Main = TakeI(.Columns.Count)\n    End With', TAKEI)).toEqual([]);
	});
});
