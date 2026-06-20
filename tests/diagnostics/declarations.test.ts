// Diagnostics tests: declarations rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import {
	analyzeProjectModule,
	projectClassMembers,
	visibleProjectNonTypeNames,
	visibleProjectTypes,
} from './helpers';

describe('analyzeModule - invalid procedure header', () => {
	it('flags a procedure name that contains a space', () => {
		const src = 'Sub My Sub\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-proc-header');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Sub');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a junk token after a Function name', () => {
		const src = 'Function Calc Total() As Long\nEnd Function\n';
		const hits = byCode(analyzeModule(src), 'invalid-proc-header');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Total');
	});

	it('does not flag a valid parameterless Sub', () => {
		const src = 'Sub Run\n    Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a valid Sub with parameters', () => {
		const src = 'Sub Greet(ByVal who As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a Function with a return type', () => {
		const src = 'Function Total() As Long\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a Property Get with a return type', () => {
		const src = 'Property Get Name() As String\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid identifier starts', () => {
	it('flags digit-start variable and parameter names', () => {
		const src =
			'Sub T(ByVal 2bad As Long)\n' +
			'    Dim 1bad As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-identifier-start');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['2bad', '1bad']);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('flags digit-start module, procedure, and compiler-constant declaration names', () => {
		const src =
			'#Const 1debug = True\n' +
			'Public Declare Sub 9Sleep Lib "kernel32" ()\n' +
			'Public Type 3Thing\n' +
			'    4Field As Long\n' +
			'End Type\n' +
			'Public Enum 5Choice\n' +
			'    6First = 1\n' +
			'End Enum\n' +
			'Sub 7Run()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-identifier-start');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'1debug',
			'9Sleep',
			'3Thing',
			'4Field',
			'5Choice',
			'6First',
			'7Run',
		]);
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('accepts ordinary and bracketed foreign names', () => {
		const src =
			'Sub T(ByVal value1 As Long)\n' +
			'    Dim value2 As Long\n' +
			'    Dim [1bad] As Long\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'invalid-identifier-start')).toHaveLength(0);
	});
});

describe('analyzeModule - module declarations inside procedures', () => {
	it('flags module-only declarations inside procedure bodies', () => {
		const src =
			'Sub T()\n' +
			'    Option Explicit\n' +
			'    Attribute T.VB_Description = "bad placement"\n' +
			'    DefLng A-Z\n' +
			'    Public leakedValue As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-in-procedure');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Option',
			'Attribute',
			'DefLng',
			'Public',
		]);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('flags unindented member Attribute lines when the module is not exported metadata source', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Debug.Print "body"\n' +
			'Attribute T.VB_Description = "bad placement"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		const hits = byCode(diagnostics, 'module-declaration-in-procedure');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Attribute');
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('does not cascade Option Explicit diagnostics on invalid first-line Attribute statements', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub Combined049AttributeInsideProc()\n' +
			'Attribute OtherProc.VB_Description = "bad placement"\n' +
			'    Debug.Print "body"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		const hits = byCode(diagnostics, 'module-declaration-in-procedure');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Attribute');
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts unindented exported member Attribute lines in the member metadata slot', () => {
		const src =
			'Public Property Get NewEnum() As IUnknown\n' +
			'Attribute NewEnum.VB_UserMemId = -4\n' +
			'    Set NewEnum = Nothing\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('flags unindented member Attribute lines after executable procedure body statements', () => {
		const src =
			'Attribute VB_Name = "Module1"\n' +
			'Sub T()\n' +
			'    Debug.Print "body"\n' +
			'Attribute T.VB_Description = "exported metadata"\n' +
			'End Sub\n';

		const hits = byCode(analyzeModule(src), 'module-declaration-in-procedure');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Attribute');
	});

	it('flags indented Attribute lines after executable procedure body statements', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "body"\n' +
			'    Attribute T.VB_Description = "bad placement"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-in-procedure');

		expectDiagnostic(src, hits, 'module-declaration-in-procedure', { span: 'Attribute' });
	});

	it('checks nested procedure blocks and ignores inactive conditional branches', () => {
		const src =
			'Sub T()\n' +
			'    If True Then\n' +
			'        DefStr A-Z\n' +
			'    End If\n' +
			'    #If Enabled Then\n' +
			'        DefLng A-Z\n' +
			'    #End If\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { conditionalCompilation: { compilerConstants: { Enabled: false } } }),
			'module-declaration-in-procedure',
		);

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DefStr']);
	});

	it('does not crash on a comment-only line after a conditional directive in a procedure', () => {
		// Regression: the first statement after a #If/#Else/#End If inside a
		// procedure body is checked against isAlternativeProcedureHeaderStatement,
		// which tokenizes the statement. A comment-only line yields zero
		// significant tokens; the header probe must not throw on the empty token
		// list (it used to, collapsing analyzeModule's catch-all to []). The
		// module reports nothing for this construct, but it must report it as an
		// empty list rather than swallowing every other rule's findings.
		const src =
			'Sub Trigger()\n' +
			'#If Win64 Then\n' +
			"    '\n" +                 // bare comment, first line after #If ... Then
			'    x = 1\n' +
			'#Else\n' +
			"    'another comment\n" +  // first line after #Else
			'    x = 2\n' +
			'#End If\n' +
			'End Sub\n' +
			'\n' +
			'Sub Dup()\n' +
			'End Sub\n' +
			'Sub Dup()\n' +            // a real, unrelated diagnostic in the same module
			'End Sub\n';

		// The crash collapsed the entire module to []; the duplicate-procedure
		// diagnostic surviving proves the rule pass ran to completion.
		expect(byCode(analyzeModule(src), 'duplicate-procedure').length).toBeGreaterThan(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('flags Type and Enum blocks inside procedure bodies without cascading to End Sub', () => {
		const src =
			'Public Sub T()\n' +
			'    Type TInside\n' +
			'        Value As Long\n' +
			'    End Type\n' +
			'    Enum EInside\n' +
			'        A = 1\n' +
			'    End Enum\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'module-declaration-in-procedure');

		expectDiagnostics(src, hits, 'module-declaration-in-procedure', [
			{ span: 'Type', message: 'Type blocks' },
			{ span: 'Enum', message: 'Enum blocks' },
		]);
		expect(byCode(diagnostics, 'statement-outside-procedure')).toHaveLength(0);
	});

	it('accepts module declarations at module level and procedure-local declarations', () => {
		const src =
			'Option Explicit\n' +
			'DefLng A-Z\n' +
			'Private moduleValue As Long\n' +
			'Sub T()\n' +
			'    Static localValue As Long\n' +
			'    Const localConst As Long = 1\n' +
			'    Dim localDim As Long\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});
});

describe('analyzeModule - module declaration placement', () => {
	it('flags active conditional Declare statements after a procedure', () => {
		const src =
			'Public Sub Combined002DeclareAfterProc()\n' +
			'    Debug.Print "procedure before declare"\n' +
			'End Sub\n' +
			'\n' +
			'#If VBA7 Then\n' +
			'    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)\n' +
			'#Else\n' +
			'    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)\n' +
			'#End If\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-after-procedure');

		expectDiagnostic(src, hits, 'module-declaration-after-procedure', {
			severity: 'error',
			span: 'Declare',
		});
	});

	it('flags only the active #ElseIf declaration after a procedure in a valid conditional block', () => {
		const src =
			'Public Sub Earlier()\n' +
			'End Sub\n' +
			'\n' +
			'#If False Then\n' +
			'    Private Const Mode As Long = 0\n' +
			'#ElseIf True Then\n' +
			'    Private Const Mode As Long = 2\n' +
			'#Else\n' +
			'    Private Const Mode As Long = 1\n' +
			'#End If\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'module-declaration-after-procedure');

		expectDiagnostic(src, hits, 'module-declaration-after-procedure', { span: 'Const' });
		expect(byCode(diagnostics, 'else-branch-order')).toHaveLength(0);
	});

	it('explains active conditional branches in declaration-after-procedure diagnostics', () => {
		const src =
			'Public Sub Earlier()\n' +
			'End Sub\n' +
			'\n' +
			'#If False Then\n' +
			'    Private Const Mode As Long = 0\n' +
			'#ElseIf True Then\n' +
			'    Private Const Mode As Long = 2\n' +
			'#Else\n' +
			'    Private Const Mode As Long = 1\n' +
			'#End If\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-after-procedure');

		expectDiagnostic(src, hits, 'module-declaration-after-procedure', { span: 'Const' });
	});

	it('suppresses declaration-order fallout inside malformed conditional-compilation blocks', () => {
		const src =
			'Public Sub Earlier()\n' +
			'End Sub\n' +
			'\n' +
			'#If False Then\n' +
			'    Private Const Mode As Long = 0\n' +
			'#Else\n' +
			'    Private Const Mode As Long = 1\n' +
			'#ElseIf True Then\n' +
			'    Private Const Mode As Long = 2\n' +
			'#End If\n' +
			'\n' +
			'Public Sub Combined048ElseifAfterElse()\n' +
			'    Debug.Print Mode\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'module-declaration-after-procedure')).toHaveLength(0);
		const branchOrder = byCode(analyzeModule(src), 'else-branch-order');
		expect(branchOrder).toHaveLength(1);
		expect(spanText(src, branchOrder[0])).toBe('#ElseIf');
	});

	it('flags module declarations after procedures and keeps procedures accepted', () => {
		const src =
			'Sub First()\nEnd Sub\n' +
			'Private Const MaxCount As Long = 1\n' +
			'Private Type RowData\n    Id As Long\nEnd Type\n' +
			'Private Enum Mode\n    ModeA = 1\nEnd Enum\n' +
			'Sub Second()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-after-procedure');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Const', 'Type', 'Enum']);
	});

	it('flags DefType statements after procedures as declaration-order errors', () => {
		const src =
			'Sub T()\n' +
			'End Sub\n' +
			'DefLng A-Z\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-after-procedure');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DefLng');
		expect(byCode(analyzeModule(src), 'statement-outside-procedure')).toHaveLength(0);
	});

	it('accepts module declarations before procedures and inactive branches after procedures', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'Private Const MaxCount As Long = 1\n' +
			'Sub T()\nEnd Sub\n' +
			'#If Enabled Then\n' +
			'    Private Declare PtrSafe Sub Hidden Lib "kernel32" ()\n' +
			'#End If\n';

		expect(byCode(
			analyzeModule(src, { conditionalCompilation: { compilerConstants: { Enabled: false } } }),
			'module-declaration-after-procedure',
		)).toHaveLength(0);
	});
});

describe('analyzeModule - module-level statements outside procedures', () => {
	it('flags a bare module-level Enum member read as a statement outside a procedure', () => {
		const src =
			'Public Enum ENeg_AmbiguousOne\n' +
			'    NegAmbiguousValue = 1\n' +
			'End Enum\n' +
			'\n' +
			'Public Enum ENeg_AmbiguousTwo\n' +
			'    NegAmbiguousValue = 2\n' +
			'End Enum\n' +
			'\n' +
			'NegAmbiguousValue\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'statement-outside-procedure');

		expectDiagnostic(src, hits, 'statement-outside-procedure', {
			severity: 'error',
			span: 'NegAmbiguousValue',
		});
		expect(byCode(diagnostics, 'ambiguous-enum-member')).toHaveLength(0);
	});

	it('flags executable module-level statements without rejecting declaration-section DefType statements', () => {
		const src =
			'Option Explicit\n' +
			'DefLng A-Z\n' +
			'Debug.Print 1\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'statement-outside-procedure');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Debug');
		expect(byCode(diagnostics, 'module-declaration-in-procedure')).toHaveLength(0);
		expect(byCode(diagnostics, 'module-declaration-after-procedure')).toHaveLength(0);
	});
});

describe('analyzeModule - reserved declaration names', () => {
	it('flags reserved keywords used as procedure and variable names', () => {
		const src =
			'Function Dim() As String\n' +
			'    Dim In As String\n' +
			'    Dim = "ok"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'invalid-declaration-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Dim', 'In']);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('flags reserved-for-implementation-use names used as declaration names', () => {
		const src =
			'Private Attribute As Long\n' +
			'Public Sub T(ByVal VB_Name As String)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-declaration-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Attribute', 'VB_Name']);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('flags reserved keywords used as type, enum, field, member, and parameter names', () => {
		const src =
			'Public Type Type\n' +
			'    For As String\n' +
			'End Type\n' +
			'Public Enum Enum\n' +
			'    In\n' +
			'End Enum\n' +
			'Public Sub T(ByVal New As String)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-declaration-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Type',
			'For',
			'Enum',
			'In',
			'New',
		]);
	});

	it('allows bracketed reserved words as foreign names', () => {
		const src =
			'Function [Dim]() As String\n' +
			'    Dim [In] As String\n' +
			'    [Dim] = [In]\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'invalid-declaration-name')).toHaveLength(0);
	});

	it('accepts Event declarations with parameter lists in object modules', () => {
		const src =
			'Public Event BeforeAdd(ByRef arr As Variant, ByRef cancel As Boolean)\n' +
			'Public Event AfterAdd(ByRef arr As Variant)\n';

		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			const diagnostics = analyzeModule(src, { moduleKind });
			expect(byCode(diagnostics, 'invalid-declaration-name')).toHaveLength(0);
			expect(byCode(diagnostics, 'unexpected-declaration-token')).toHaveLength(0);
			expect(byCode(diagnostics, 'duplicate-module-variable')).toHaveLength(0);
			expect(byCode(diagnostics, 'event-declaration-module-kind')).toHaveLength(0);
		}
	});

	it('accepts user-defined type fields named Type', () => {
		const src =
			'Private Type PICTDESC\n' +
			'    size As Long\n' +
			'    Type As Long\n' +
			'End Type\n';

		expect(byCode(analyzeModule(src), 'invalid-declaration-name')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('allows runtime and host-global shadowing declarations while rejecting reserved declaration names', () => {
		const src =
			'Private Left As Long\n' +
			'Private Right As Long\n' +
			'Private Date As Variant\n' +
			'Private Collection As Object\n' +
			'Private Application As Object\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'invalid-declaration-name');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Date');
		expect(byCode(diagnostics, 'invalid-as-type-name')).toHaveLength(0);
		expect(byCode(diagnostics, 'statement-outside-procedure')).toHaveLength(0);
	});
});

describe('analyzeModule - As type name validation', () => {
	it('flags runtime functions used as declaration type names', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest3 As Int\n' +
			'    shouldErrorTest3 = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-as-type-name');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('Int');
	});

	it('uses project and host type resolution across type-name positions', () => {
		const src =
			'Public Type Payload\n' +
			'    Owner As Person\n' +
			'End Type\n' +
			'\n' +
			'Public Function Make(ByVal state As Status, ByVal sheet As Worksheet) As Person\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    If TypeOf p Is Person Then Debug.Print "ok"\n' +
			'    Set Make = p\n' +
			'End Function\n';
		const modules = [
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Types', source: 'Public Enum Status\n    Active\nEnd Enum\n' },
		];

		expect(
			byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-as-type-name'),
		).toHaveLength(0);
	});

	it('accepts OLE Automation interfaces even when project enum members share the name', () => {
		const src =
			'Public Enum EKnownIID\n' +
			'    IUnknown\n' +
			'End Enum\n' +
			'Private Declare PtrSafe Function RegisterActiveObject Lib "oleaut32" (ByVal pUnk As IUnknown) As Long\n' +
			'Private Declare PtrSafe Function RegisterQualified Lib "oleaut32" (ByVal pUnk As stdole.IUnknown) As Long\n';
		expect(byCode(analyzeModule(src), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('uses the default VBA7 branch for project types defined in paired branches', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Payload\n' +
			'End Sub\n';
		const types =
			'#If VBA7 Then\n' +
			'Public Type Payload\n' +
			'    Id As LongPtr\n' +
			'End Type\n' +
			'#Else\n' +
			'Public Type Payload\n' +
			'    Id As Long\n' +
			'End Type\n' +
			'#End If\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Types', source: types },
		];

		const diagnostics = analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		});
		expect(byCode(diagnostics, 'invalid-as-type-name')).toHaveLength(0);
	});

	it('accepts qualified visible project type names in declaration positions', () => {
		const src =
			'Public Type Payload\n' +
			'    Location As Geometry.TPoint\n' +
			'End Type\n' +
			'\n' +
			'Public Sub T(ByVal state As Workflow.Status)\n' +
			'    Dim p As Geometry.TPoint\n' +
			'End Sub\n';
		const modules = [
			{
				moduleName: 'Geometry',
				source: 'Public Type TPoint\n    X As Long\nEnd Type\n',
			},
			{
				moduleName: 'Workflow',
				source: 'Public Enum Status\n    Active\nEnd Enum\n',
			},
			{
				moduleName: 'OtherGeometry',
				source: 'Public Type TPoint\n    Y As Long\nEnd Type\n',
			},
		];

		expect(
			byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-as-type-name'),
		).toHaveLength(0);
	});

	it('accepts creatable project classes and UserForms after New', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As New Person\n' +
			'    Dim f As New CustomerForm\n' +
			'    Dim value As Object\n' +
			'    Set value = New Person\n' +
			'    Set value = New CustomerForm\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'CustomerForm', moduleKind: 'userform' as const, source: '' },
		];

		expect(
			byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-new-type-name'),
		).toHaveLength(0);
	});

	it('flags resolved source-backed non-creatable types after New while deferring host types', () => {
		const src =
			'Public Type Payload\n' +
			'    Id As Long\n' +
			'End Type\n' +
			'\n' +
			'Public Enum Status\n' +
			'    Active\n' +
			'End Enum\n' +
			'\n' +
			'Public Sub T()\n' +
			'    Dim eagerSheet As New Worksheet\n' +
			'    Dim eagerLong As New Long\n' +
			'    Dim value As Object\n' +
			'    Set value = New Status\n' +
			'    Set value = New Payload\n' +
			'    Set value = New Sheet1\n' +
			'    Set value = New Worksheet\n' +
			'    Set value = New Long\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Sheet1', moduleKind: 'document' as const, source: '' },
		];

		const diags = analyzeProjectModule(src, modules, 'Consumer');
		const hits = byCode(diags, 'invalid-new-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Long',
			'Status',
			'Payload',
			'Sheet1',
			'Long',
		]);
		expect(byCode(diags, 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags qualified non-creatable project types after New', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim value As Object\n' +
			'    Set value = New Geometry.TPoint\n' +
			'    Set value = New Workflow.Status\n' +
			'End Sub\n';
		const modules = [
			{
				moduleName: 'Geometry',
				source: 'Public Type TPoint\n    X As Long\nEnd Type\n',
			},
			{
				moduleName: 'Workflow',
				source: 'Public Enum Status\n    Active\nEnd Enum\n',
			},
		];

		const hits = byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-new-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['TPoint', 'Status']);
	});

	it('defers unresolved New type names to the project-wide binder and external references', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim eager As New ExternalThing\n' +
			'    Dim value As Object\n' +
			'    Set value = New OtherExternalThing\n' +
			'End Sub\n';
		const diags = analyzeModule(src);

		expect(byCode(diags, 'invalid-new-type-name')).toHaveLength(0);
		expect(byCode(diags, 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags runtime functions used in shared type-name positions', () => {
		const src =
			'Public Type Bag\n' +
			'    Field As Left\n' +
			'End Type\n' +
			'\n' +
			'Public Function Make(ByVal item As Int) As Right\n' +
			'    Dim value As Mid\n' +
			'    Set value = New Left\n' +
			'    If TypeOf value Is Right Then Debug.Print "bad"\n' +
			'End Function\n';

		const hits = byCode(analyzeModule(src), 'invalid-as-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Left',
			'Int',
			'Right',
			'Mid',
			'Left',
			'Right',
		]);
	});

	it('lets a project type shadow a runtime function name', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim value As Int\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Module1', source: src },
			{ moduleName: 'Int', moduleKind: 'class' as const, source: '' },
		];

		expect(byCode(analyzeModule(src, {
			moduleName: 'Module1',
			projectTypes: visibleProjectTypes(modules, 'Module1'),
		}), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags reserved identifiers used in shared type-name positions', () => {
		const src =
			'Implements Return\n' +
			'Public Sub T()\n' +
			'    Dim value As Dim\n' +
			'    Set value = New For\n' +
			'    If TypeOf value Is In Then Debug.Print "bad"\n' +
			'End Sub\n';

		const hits = byCode(analyzeModule(src), 'invalid-as-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Return',
			'Dim',
			'For',
			'In',
		]);
	});

	it('flags visible project declarations that are known not to be types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As DoWork\n' +
			'    Dim b As SharedValue\n' +
			'    Dim c As Active\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{
				moduleName: 'Helpers',
				source:
					'Public Sub DoWork()\nEnd Sub\n' +
					'Public SharedValue As Long\n' +
					'Public Enum Status\n    Active\nEnd Enum\n',
			},
		];

		const hits = byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'DoWork',
			'SharedValue',
			'Active',
		]);
	});

	it('lets a project type shadow a visible non-type declaration name', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim value As Person\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Helpers', source: 'Public Sub Person()\nEnd Sub\n' },
		];

		expect(byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags ambiguous visible project type names', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim state As Status\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Status', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Types', source: 'Public Enum Status\n    Active\nEnd Enum\n' },
		];

		const hits = byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name');

		expectDiagnostic(src, hits, 'invalid-as-type-name', { span: 'Status' });
	});

	it('flags duplicate visible project type names even when their kind matches', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Payload\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'TypesA', source: 'Public Type Payload\n    Id As Long\nEnd Type\n' },
			{ moduleName: 'TypesB', source: 'Public Type Payload\n    Name As String\nEnd Type\n' },
		];

		const hits = byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name');

		expectDiagnostic(src, hits, 'invalid-as-type-name', { span: 'Payload' });
	});

	it('defers broad unknown type names to the project-wide binder', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest2 As Intege\n' +
			'    shouldErrorTest2 = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-as-type-name')).toHaveLength(0);
	});
});

describe('analyzeModule - declaration initializer', () => {
	it('flags a module-level Dim with a VB.NET-style initializer', () => {
		const src = 'Private x As Long = 1\n';
		const hits = byCode(analyzeModule(src), 'dim-initializer');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('=');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a local Dim with an initializer', () => {
		const src = 'Sub T()\n    Dim n As Long = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(1);
	});

	it('flags an initializer without an As clause', () => {
		const src = 'Sub T()\n    Dim n = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(1);
	});

	it('does not flag a Const declaration', () => {
		const src = 'Const Pi As Double = 3.14\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});

	it('does not flag a plain declaration', () => {
		const src = 'Sub T()\n    Dim n As Long\n    n = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});

	it('does not flag an array bound that contains digits', () => {
		const src = 'Sub T()\n    Dim a(1 To 10) As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});
});

describe('analyzeModule - type-declaration suffixes', () => {
	it('flags rejecting declarations that combine a type-declaration character with an As clause', () => {
		const src =
			'Private Type Header\n' +
			'    Code$ As String\n' +
			'End Type\n' +
			'\n' +
			'Public Function GetName$() As String\n' +
			'    GetName = "XLIDE"\n' +
			'End Function\n' +
			'\n' +
			'Public Sub Demo(ByVal label$ As String)\n' +
			'    Const answer% As Long = 1\n' +
			'    Dim value$ As Long\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'type-declaration-character-as-clause');

		expectDiagnostics(src, hits, 'type-declaration-character-as-clause', [
			{ span: '$' },
			{ span: '$' },
			{ span: '$' },
			{ span: '%' },
			{ span: '$' },
		]);
		expect(byCode(diagnostics, 'invalid-procedure-header')).toHaveLength(0);
		expect(byCode(diagnostics, 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('accepts VBE-verified suffix controls', () => {
		const src =
			'Private Type Header\n' +
			'    Code$\n' +
			'End Type\n' +
			'\n' +
			'Public Property Get Name$() As String\n' +
			'    Name = "XLIDE"\n' +
			'End Property\n' +
			'\n' +
			'Public Function GetName$()\n' +
			'    GetName = "XLIDE"\n' +
			'End Function\n' +
			'\n' +
			'Public Sub Demo(ByVal label$)\n' +
			'    Const answer% = 1\n' +
			'    Dim total&, name$, price@, ratio#, flag%\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'type-declaration-character-as-clause')).toHaveLength(0);
	});
});

describe('analyzeModule - unexpected declaration tokens', () => {
	it('flags a bare identifier after a complete local As type', () => {
		const src = 'Sub T()\n    Dim s1 As String thisshoulderror\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'unexpected-declaration-token', {
			severity: 'error',
			span: 'thisshoulderror',
		});
	});

	it('flags trailing tokens in module declarations, parameters, and Type fields', () => {
		const src =
			'Private moduleName As String junk\n' +
			'Private Type Customer\n' +
			'    Name As String extra\n' +
			'End Type\n' +
			'Sub T(ByVal label As String trailing)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unexpected-declaration-token');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'junk',
			'extra',
			'trailing',
		]);
	});

	it('accepts normal declaration separators and parameter defaults', () => {
		const src =
			'Private moduleName As String\n' +
			'Sub T(Optional ByVal label As String = "ok")\n' +
			'    Dim first As String, second As Long: Dim third As Variant\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('accepts fixed-length string declarations and qualified type names', () => {
		const src =
			'Private fixedName As String * 10\n' +
			'Private Type Header\n' +
			'    Code As String * 4\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Dim localName As String * 20\n' +
			'    Dim workbook As Excel.Workbook\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('flags extra tokens after a fixed-length string suffix', () => {
		const src = 'Sub T()\n    Dim fixedName As String * 10 junk\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unexpected-declaration-token');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('junk');
	});
});

describe('analyzeModule - fixed-length String bounds', () => {
	it('accepts verified literal boundaries and literal Const lengths', () => {
		const src =
			'Private Const HeaderCodeLength As Long = 65526\n' +
			'Private moduleName As String * 1\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const MaxNameLength As Long = 20\n' +
			'    Dim localName As String * MaxNameLength\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-length-string-size')).toHaveLength(0);
	});

	it('resolves compound integer Const expressions used as fixed-length String sizes', () => {
		const src =
			'Private Const BaseLength As Long = 32763\n' +
			'Private Const HeaderCodeLength As Long = BaseLength * 2 + 1\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const LocalBase As Long = 10\n' +
			'    Const MaxNameLength As Long = LocalBase + 10\n' +
			'    Const LocalTooSmall As Long = LocalBase - 10\n' +
			'    Dim localName As String * MaxNameLength\n' +
			'    Dim badLocalName As String * LocalTooSmall\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'fixed-length-string-size', [
			{ span: 'HeaderCodeLength', message: 'got 65527' },
			{ span: 'LocalTooSmall', message: 'got 0' },
		]);
	});

	it('resolves non-decimal integer Consts used as fixed-length String sizes', () => {
		const src =
			'Private Const HeaderCodeLength As Long = &H14\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const LocalTooSmall As Long = &O0\n' +
			'    Dim localName As String * LocalTooSmall\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'fixed-length-string-size', {
			span: 'LocalTooSmall',
		});
	});

	it('defers unknown and non-deterministic Const length expressions', () => {
		const src =
			'Private Const FractionLength As Long = 20 / 2\n' +
			'Private moduleName As String * MissingLength\n' +
			'Private Type Header\n' +
			'    Code As String * FractionLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const RuntimeLength As Long = CLng(20)\n' +
			'    Dim localName As String * RuntimeLength\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-length-string-size')).toHaveLength(0);
	});

	it('flags literal lengths outside the VBE-verified range everywhere declarations are parsed', () => {
		const src =
			'Private moduleName As String * 0\n' +
			'Private Type Header\n' +
			'    Code As String * 65527\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    If True Then\n' +
			'        Dim localName As String * 0\n' +
			'    End If\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'fixed-length-string-size', [
			{ span: '0' },
			{ span: '65527' },
			{ span: '0' },
		]);
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('flags simple Const lengths outside the VBE-verified range', () => {
		const src =
			'Private Const ModuleTooLong As Long = 65527\n' +
			'Private moduleName As String * ModuleTooLong\n' +
			'Sub T()\n' +
			'    Const LocalTooSmall As Long = 0\n' +
			'    Dim localName As String * LocalTooSmall\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'fixed-length-string-size', [
			{ span: 'ModuleTooLong', message: 'got 65527' },
			{ span: 'LocalTooSmall', message: 'got 0' },
		]);
	});

	it('lets procedure-local Const lengths shadow module Const lengths', () => {
		const src =
			'Private Const NameLength As Long = 65527\n' +
			'Sub T()\n' +
			'    Const NameLength As Long = 20\n' +
			'    Dim localName As String * NameLength\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-length-string-size')).toHaveLength(0);
	});

	it('resolves bracketed Const names used as fixed-length String sizes', () => {
		const src =
			'Private Const [Name Length] As Long = 0\n' +
			'Sub T()\n' +
			'    Dim localName As String * [Name Length]\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('[Name Length]');
	});
});

describe('analyzeModule - parameter order', () => {
	it('flags a required parameter after an Optional one', () => {
		const src =
			'Sub T(Optional ByVal a As Long = 1, ByVal b As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'required-param-after-optional');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('b');
	});

	it('accepts trailing Optional parameters', () => {
		const src =
			'Sub T(ByVal a As Long, Optional ByVal b As Long = 1)\nEnd Sub\n';
		expect(
			byCode(analyzeModule(src), 'required-param-after-optional'),
		).toHaveLength(0);
	});

	it('flags array parameter parentheses after the As type and suppresses the generic token error', () => {
		const src =
			'Public Sub NegParam06_BadArrayParameterSyntax(ByVal values As Long())\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'parameter-array-as-type-syntax');

		expectDiagnostic(src, hits, 'parameter-array-as-type-syntax', {
			severity: 'error',
			span: '()',
		});
		expect(byCode(diagnostics, 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('accepts array parameter parentheses after the parameter name', () => {
		const src = 'Sub T(ByVal values() As Long)\nEnd Sub\n';

		expect(byCode(analyzeModule(src), 'parameter-array-as-type-syntax')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('flags a ParamArray that is not last', () => {
		const src = 'Sub T(ParamArray items() As Variant, ByVal n As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-not-last');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('items');
	});

	it('flags a ParamArray combined with Optional parameters', () => {
		const src =
			'Public Sub Combined011ParamarrayWithOptional(Optional ByVal prefix As String = "x", ParamArray values() As Variant)\n' +
			'    Debug.Print prefix\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-with-optional');

		expectDiagnostic(src, hits, 'paramarray-with-optional', { severity: 'error', span: 'values' });
	});

	it('flags a ParamArray typed as a non-Variant element type', () => {
		const src =
			'Public Sub Combined012TypedParamarray(ParamArray values() As String)\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-non-variant');

		expectDiagnostic(src, hits, 'paramarray-non-variant', { severity: 'error', span: 'values' });
	});

	it('accepts a trailing ParamArray', () => {
		const src =
			'Sub T(ByVal n As Long, ParamArray items() As Variant)\nEnd Sub\n' +
			'Sub U(ParamArray values())\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'paramarray-not-last')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'paramarray-with-optional')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'paramarray-non-variant')).toHaveLength(0);
	});
});

describe('analyzeModule - property setter shape', () => {
	it('flags Property Let declarations with no value parameter', () => {
		const src =
			'Public Property Let Name()\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-setter-missing-value');

		expectDiagnostic(src, hits, 'property-setter-missing-value', {
			severity: 'error',
			span: 'Name',
		});
	});

	it('flags Property Set declarations with no value parameter', () => {
		const src =
			'Public Property Set Child()\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-setter-missing-value');

		expectDiagnostic(src, hits, 'property-setter-missing-value', { span: 'Child' });
	});

	it('flags Property Set declarations with scalar value parameters', () => {
		const src =
			'Public Property Set Number(ByVal value As Long)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-set-scalar-value');

		expectDiagnostic(src, hits, 'property-set-scalar-value', { severity: 'error', span: 'value' });
	});

	it('flags Property Let declarations with Object value parameters', () => {
		const src =
			'Public Property Let NegProp06_LetObjectValue(ByVal Value As Object)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-let-object-value');

		expectDiagnostic(src, hits, 'property-let-object-value', { severity: 'error', span: 'Value' });
	});

	it('flags Property Let declarations with known host and project object value parameters', () => {
		const src =
			'Public Property Let Sheet(ByVal value As Worksheet)\n' +
			'End Property\n' +
			'Public Property Let Customer(ByVal assigned As Person)\n' +
			'End Property\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: [
					{ name: 'Person', kind: 'class', moduleName: 'Person', members: [] },
				],
			}),
			'property-let-object-value',
		);

		expectDiagnostics(src, hits, 'property-let-object-value', [
			{ span: 'value', message: 'Worksheet' },
			{ span: 'assigned', message: 'Person' },
		]);
	});

	it('flags Property Let and Set declarations with return types', () => {
		const src =
			'Public Property Let NegProp02_LetWithReturnType(ByVal value As Long) As Long\n' +
			'End Property\n' +
			'Public Property Set Child(ByVal value As Object) As Object\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-setter-return-type');

		expect(hits).toHaveLength(2);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
		expectDiagnostics(src, hits, 'property-setter-return-type', [
			{ span: 'As Long', message: 'Property Let' },
			{ span: 'As Object', message: 'Property Set' },
		]);
	});

	it('flags Property Let declarations missing indexed getter parameters', () => {
		const src =
			'Public Property Get Item(ByVal index As Long) As String\n' +
			'End Property\n' +
			'Public Property Let Item(ByVal value As String)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-accessor-signature-mismatch');

		expectDiagnostic(src, hits, 'property-accessor-signature-mismatch', {
			severity: 'error',
			span: 'Item',
		});
	});

	it('flags Property Let and Set indexed parameter shape mismatches', () => {
		const src =
			'Public Property Get Label(ByVal index As Long) As String\n' +
			'End Property\n' +
			'Public Property Let Label(ByVal index As String, ByVal value As String)\n' +
			'End Property\n' +
			'Public Property Get Child(ByVal index As Long) As Object\n' +
			'End Property\n' +
			'Public Property Set Child(ByRef index As Long, ByVal value As Object)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-accessor-signature-mismatch');

		expectDiagnostics(src, hits, 'property-accessor-signature-mismatch', [
			{ span: 'Label', message: 'type must match' },
			{ span: 'Child', message: 'passing mode' },
		]);
	});

	it('accepts standalone read-only indexed Property Get declarations', () => {
		const src =
			'Public Property Get NegProp07_IndexedGet(ByVal index As Long) As Long\n' +
			'    NegProp07_IndexedGet = index\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'property-accessor-signature-mismatch')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'property-setter-missing-value')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'function-return-not-assigned')).toHaveLength(0);
	});

	it('accepts Property Get and object-shaped setters with value parameters', () => {
		const src =
			'Public Property Get Name() As String\n' +
			'End Property\n' +
			'Public Property Let Name(ByVal value As String)\n' +
			'End Property\n' +
			'Public Property Set Child(ByVal value As Object)\n' +
			'End Property\n' +
			'Public Property Set IndexedChild(ByVal index As Long, ByVal value As Person)\n' +
			'End Property\n';
		expect(byCode(analyzeModule(src), 'property-setter-missing-value')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'property-set-scalar-value')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'property-let-object-value')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'property-accessor-signature-mismatch')).toHaveLength(0);
	});

	it('accepts write-only setters and inactive mismatched property accessors', () => {
		const src =
			'Public Property Let Name(ByVal value As String)\n' +
			'End Property\n' +
			'#If Win64 Then\n' +
			'Public Property Get Item(ByVal index As Long) As String\n' +
			'End Property\n' +
			'#Else\n' +
			'Public Property Let Item(ByVal value As String)\n' +
			'End Property\n' +
			'#End If\n';
		expect(
			byCode(
				analyzeModule(src, { conditionalCompilation: { Win64: true } }),
				'property-accessor-signature-mismatch',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - parameter default values', () => {
	it('flags nonnumeric string defaults for numeric and Boolean Optional parameters', () => {
		const src =
			'Sub T(Optional ByVal count As Long = "bad", Optional ByVal enabled As Boolean = "bad")\nEnd Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'parameter-default-type-mismatch', [
			{ span: '"bad"', message: ['count', 'expects Long'] },
			{ span: '"bad"', message: ['enabled', 'expects Boolean'] },
		]);
	});

	it('flags non-Nothing Optional object parameter defaults', () => {
		const src =
			'Public Sub NegParam07_OptionalObjectDefaultNonNothing(Optional ByVal obj As Object = 1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'parameter-default-type-mismatch');

		expectDiagnostic(src, hits, 'parameter-default-type-mismatch', {
			severity: 'error',
			span: '1',
		});
	});

	it('flags scalar defaults on Optional array parameters', () => {
		const src =
			'Public Sub NegParam08_OptionalArrayDefault(Optional values() As Long = 0)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'parameter-default-type-mismatch');

		expectDiagnostic(src, hits, 'parameter-default-type-mismatch', {
			severity: 'error',
			span: '0',
		});
	});

	it('flags non-Nothing defaults for known host and project object parameters', () => {
		const src =
			'Sub T(Optional ByVal sheet As Worksheet = "Sheet1", Optional ByVal person As Person = True)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: [
					{ name: 'Person', kind: 'class', moduleName: 'Person', members: [] },
				],
			}),
			'parameter-default-type-mismatch',
		);

		expectDiagnostics(src, hits, 'parameter-default-type-mismatch', [
			{ span: '"Sheet1"', message: 'Nothing' },
			{ span: 'True', message: 'Nothing' },
		]);
	});

	it('accepts oracle-backed scalar Optional default controls', () => {
		const src =
			'Sub T(Optional ByVal count As Long = 1, Optional ByVal fromText As Long = "1", Optional ByVal label As String = "ok", Optional ByVal enabled As Boolean = True, Optional ByVal obj As Object = Nothing, Optional ByVal sheet As Worksheet = Nothing, Optional values() As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'parameter-default-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - Option placement', () => {
	it('flags an Option after a declaration', () => {
		const src = 'Private m_Count As Long\nOption Explicit\n';
		const hits = byCode(analyzeModule(src), 'option-after-declaration');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('Option');
	});

	it('flags an Option after a procedure', () => {
		const src = 'Sub T()\nEnd Sub\nOption Base 1\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(1);
	});

	it('accepts Options at the top of the module', () => {
		const src = 'Option Explicit\nOption Base 1\nPrivate m_Count As Long\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(0);
	});

	it('accepts Options after Attribute lines', () => {
		const src =
			'Attribute VB_Name = "Module1"\nOption Explicit\nSub T()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(0);
	});
});

describe('analyzeModule - declaration no-diagnostic boundary controls (rule audit backfill)', () => {
	it('stays quiet for a digit-start procedure name owned by invalid-identifier-start', () => {
		// invalid-proc-header defers the malformed digit-start header to
		// invalid-identifier-start (which owns the precise declaration-name range).
		const src = 'Sub 1Bad()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('stays quiet for Property Set value parameters with unresolved or Variant types', () => {
		// property-set-scalar-value fires only for a provably-scalar final value
		// param; an unresolved type name and As Variant are not known scalars.
		const src =
			'Public Property Set Gadget(ByVal value As Widget)\n' +
			'End Property\n' +
			'Public Property Set Anything(ByVal value As Variant)\n' +
			'End Property\n';
		expect(byCode(analyzeModule(src), 'property-set-scalar-value')).toHaveLength(0);
	});

	it('stays quiet for a Property Let whose value parameter has an unresolved type name', () => {
		// property-let-object-value fires only when the value param resolves to a
		// known object type; an unresolved name (no projectClassMembers entry) stays quiet.
		const src =
			'Public Property Let Customer(ByVal assigned As Person)\n' +
			'End Property\n';
		expect(byCode(analyzeModule(src), 'property-let-object-value')).toHaveLength(0);
	});
});
