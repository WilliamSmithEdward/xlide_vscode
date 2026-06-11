import { describe, it, expect } from 'vitest';
import {
	analyzeModule,
	DIAGNOSTIC_RULES,
	STRUCTURAL_DIAGNOSTIC_RULES,
	diagnosticMetadataForCode,
	diagnosticSourceForCode,
	isXlideDiagnosticSource,
	type HostObjectModel,
} from '../src/analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
} from '../src/vbaProjectAnalysis';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from './helpers/diagnostics';
import {
	analyzeProjectModule,
	projectClassMembers,
	projectMemberSurfaces,
	projectProcedures,
	visibleProjectNonTypeNames,
	visibleProjectTypes,
	type ProjectTestModule,
} from './diagnostics/helpers';

describe('vbaProjectAnalysis helper', () => {
	it('derives project member surfaces from the live current-module overlay', () => {
		const savedClass = [
			'Public Property Get Name() As String',
			'End Property',
		].join('\n');
		const liveClass = [
			'Public Property Get Age() As Long',
			'End Property',
		].join('\n');
		const project = buildVbaProjectIndex(
			[
				{ moduleName: 'Person', moduleKind: 'class', source: savedClass },
				{ moduleName: 'Caller', moduleKind: 'standard', source: '' },
			],
			{ moduleName: 'Person', moduleKind: 'class', source: liveClass },
		);
		const person = projectAnalysisOptionsForModule(project, 'Caller')
			.projectClassMembers
			?.find((surface) => surface.name === 'Person');

		expect(person?.members.map((member) => member.name)).toEqual(['Age']);
	});
});

describe('analyzeModule - assignment to constant', () => {
	it('flags assigning to a module-level Const', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    MAX = 5\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'const-assignment');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MAX');
	});

	it('flags assigning to a local Const', () => {
		const src = 'Sub T()\n    Const PI As Double = 3.14\n    PI = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(1);
	});

	it('uses local shadows before module-level Const declarations', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n' +
			'    Dim MAX As Long\n' +
			'    MAX = 5\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('flags assigning to a visible exported Const', () => {
		const caller = 'Sub T()\n    SharedMax = 5\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source: 'Public Const SharedMax As Long = 10\n',
			},
		], 'Caller');
		const hits = byCode(diagnostics, 'const-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedMax');
	});

	it('keeps ambiguous exported Const assignments quiet', () => {
		const caller = 'Sub T()\n    SharedMax = 5\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'GlobalsA',
				source: 'Public Const SharedMax As Long = 10\n',
			},
			{
				moduleName: 'GlobalsB',
				source: 'Public Const SharedMax As Long = 20\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'const-assignment')).toHaveLength(0);
	});

	it('does not flag comparing a Const in a condition', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    If MAX = 10 Then Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('does not flag assigning to a non-constant variable', () => {
		const src = 'Sub T()\n    Dim x As Long\n    x = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('does not flag a member or indexed left-hand side that shares a Const name', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    obj.MAX = 5\n    arr(MAX) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('flags assigning to a Const after a numeric line label', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n' +
			'10 MAX = 5\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'const-assignment');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MAX');
	});
});

describe('analyzeModule - general contract', () => {
	it('every diagnostic rule declares category, VBE equivalence, and evidence kind', () => {
		const categories = new Set([
			'syntax',
			'lexer',
			'parser',
			'realtime-recovery',
			'declaration',
			'semantic',
			'project-symbol',
			'module-kind',
			'excel-host',
			'style',
		]);
		const evidenceKinds = new Set([
			'compile-error',
			'deterministic-runtime-error',
			'runtime-risk',
			'style-policy',
		]);
		for (const [name, rule] of Object.entries({
			...DIAGNOSTIC_RULES,
			...STRUCTURAL_DIAGNOSTIC_RULES,
		})) {
			expect(categories.has(rule.category), name).toBe(true);
			expect(typeof rule.vbeCompileEquivalent, name).toBe('boolean');
			expect(evidenceKinds.has(rule.diagnosticKind), name).toBe(true);
			if (rule.vbeCompileEquivalent) {
				expect(rule.diagnosticKind, name).toBe('compile-error');
			}
		}
	});

	it('never throws on malformed input', () => {
		const samples = ['', 'Sub', 'End Sub', 'Dim', '"', 'If Then', ':::'];
		for (const s of samples) {
			expect(() => analyzeModule(s)).not.toThrow();
		}
	});

	it('every emitted code is in the rule catalogue', () => {
		const known = new Set(
			Object.values(DIAGNOSTIC_RULES).map((r) => r.code),
		);
		const src =
			'Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n' +
			'Const C As Long = 1\nSub Bar()\n    C = 2\n    MsgBox "x\nEnd Sub\n';
		for (const d of analyzeModule(src)) {
			expect(known.has(d.code)).toBe(true);
		}
	});

	it('resolves metadata for semantic and structural diagnostic codes', () => {
		expect(diagnosticMetadataForCode('undeclared-variable')).toMatchObject({
			category: 'project-symbol',
			vbeCompileEquivalent: true,
			diagnosticKind: 'compile-error',
		});
		expect(diagnosticMetadataForCode('missing-block-closer')).toMatchObject({
			title: STRUCTURAL_DIAGNOSTIC_RULES.missingBlockCloser.title,
			category: 'syntax',
			vbeCompileEquivalent: true,
			diagnosticKind: 'compile-error',
		});
		expect(diagnosticMetadataForCode('not-a-real-rule')).toBeUndefined();
	});

	it('uses metadata source labels for Problems filtering without changing rule codes', () => {
		expect(diagnosticSourceForCode('missing-block-closer')).toBe('XLIDE/VBE');
		expect(diagnosticSourceForCode('undeclared-variable')).toBe('XLIDE/VBE');
		expect(diagnosticSourceForCode('string-arithmetic-coercion')).toBe('XLIDE/runtime');
		expect(diagnosticSourceForCode('missing-return-assignment')).toBe('XLIDE/risk');
		expect(diagnosticSourceForCode('option-explicit-missing')).toBe('XLIDE/style');
		expect(diagnosticSourceForCode(undefined)).toBe('XLIDE');
		expect(isXlideDiagnosticSource('XLIDE')).toBe(true);
		expect(isXlideDiagnosticSource('XLIDE/VBE')).toBe(true);
		expect(isXlideDiagnosticSource('typescript')).toBe(false);
	});

	it('produces a clean module with no diagnostics', () => {
		const src =
			'Option Explicit\n\n' +
			'Const MAX As Long = 10\n\n' +
			'Sub Greet(ByVal name As String)\n' +
			'    Dim msg As String\n' +
			'    msg = "Hello " & name\n' +
			'    MsgBox msg\n' +
			'End Sub\n';
		expect(analyzeModule(src)).toHaveLength(0);
	});
});

describe('analyzeModule - scalar member access', () => {
	it('flags a trailing dot on a local String variable', () => {
		const src = 'Sub Main()\n    Dim value As String\n    value.\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'scalar-member-access', {
			severity: 'error',
			span: 'value.',
		});
	});

	it('flags named member access on a local Integer variable', () => {
		const src = 'Sub Main()\n    Dim value As Integer\n    value.Length\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'scalar-member-access', { span: 'value.' });
	});

	it('flags scalar parameters and module variables', () => {
		const src =
			'Private moduleName As String\n' +
			'Sub Main(ByVal count As Long)\n' +
			'    moduleName.Length\n' +
			'    count.Value\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['moduleName.', 'count.']);
	});

	it('uses visible exported scalar globals for member receiver types', () => {
		const caller =
			'Sub Main()\n' +
			'    SharedText.Length\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'scalar-member-access',
		);

		expectDiagnostic(caller, hits, 'scalar-member-access', { span: 'SharedText.' });
	});

	it('does not leak exported scalar receiver types through local untyped shadows', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim SharedText\n' +
			'    SharedText.Length\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-member-access')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported scalar receivers quiet', () => {
		const caller =
			'Sub Main()\n' +
			'    SharedText.Length\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-member-access')).toHaveLength(0);
	});

	it('does not treat member names in qualified chains as bare scalar receivers', () => {
		const src =
			'Sub Main()\n' +
			'    Dim Value As String\n' +
			'    Dim item As Variant\n' +
			'    item.Value.Length\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(0);
	});

	it('treats colon-separated declaration and member access statements independently', () => {
		const src = 'Sub Main()\n    Dim value As String: value.Length\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
	});

	it('does not flag Variant, Object, or unknown receiver types', () => {
		const src =
			'Sub Main()\n' +
			'    Dim flexible As Variant\n' +
			'    Dim item As Object\n' +
			'    Dim person As Person\n' +
			'    flexible.\n' +
			'    item.\n' +
			'    person.\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(0);
	});
});

describe('analyzeModule - object variable not set', () => {
	it('flags straight-line local object member access before Set', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    obj.ToString\n' +
			'    Dim ws As Worksheet\n' +
			'    ws.Range("A1").Value = 1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'object-variable-not-set');

		expectDiagnostics(src, hits, 'object-variable-not-set', [
			{ severity: 'error', span: 'obj' },
			{ span: 'ws' },
		]);
	});

	it('accepts object locals after Set and flags them again after Set Nothing', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = New Collection\n' +
			'    obj.ToString\n' +
			'    Set obj = Nothing\n' +
			'    obj.ToString\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'object-variable-not-set');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('obj');
	});

	it('lets a provable missing member diagnostic supersede runtime not-set', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws.DefinitelyMissingMember\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);

		expect(byCode(diagnostics, 'object-variable-not-set')).toHaveLength(0);
		const memberHits = byCode(diagnostics, 'member-not-found');
		expect(memberHits).toHaveLength(1);
		expect(spanText(src, memberHits[0])).toBe('DefinitelyMissingMember');
	});

	it('keeps parameters module-level objects Static locals and helper-initialized locals quiet', () => {
		const src =
			'Private moduleObj As Object\n' +
			'Private Sub Initialize(ByRef target As Object)\n' +
			'End Sub\n' +
			'Public Sub T(ByVal param As Object)\n' +
			'    Static cached As Object\n' +
			'    Dim initialized As Object\n' +
			'    Initialize initialized\n' +
			'    moduleObj.ToString\n' +
			'    param.ToString\n' +
			'    cached.ToString\n' +
			'    initialized.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('stays quiet after branch-local initialization makes straight-line state unknown', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    If Ready Then\n' +
			'        Set obj = New Collection\n' +
			'    End If\n' +
			'    obj.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('ignores inactive conditional-compilation member access', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'#If Enabled Then\n' +
			'    obj.ToString\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});
});

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

describe('analyzeModule - unbalanced parentheses', () => {
	it('flags a missing closing parenthesis', () => {
		const src = 'Sub T()\n    x = (1 + 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('(');
		expect(hits[0].severity).toBe('error');
	});

	it('flags an unexpected closing parenthesis', () => {
		const src = 'Sub T()\n    x = 1 + 2)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe(')');
	});

	it('reports at most one diagnostic per statement', () => {
		const src = 'Sub T()\n    x = ((1 + 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(1);
	});

	it('does not flag balanced and nested parentheses', () => {
		const src =
			'Sub T()\n' +
			'    x = (1 + (2 * 3))\n' +
			'    Debug.Print Left("ab", 1)\n' +
			'    Cells(1, 1).Value = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('does not count parentheses inside strings or comments', () => {
		const src =
			'Sub T()\n' +
			'    x = "a (b c"  \' a ) in a comment\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('balances across a line continuation', () => {
		const src = 'Sub T()\n    x = (1 + _\n        2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('treats a colon-separated statement independently', () => {
		const src = 'Sub T()\n    a = (1 + 2) : b = 3)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe(')');
	});
});

describe('analyzeModule - argument count', () => {
	it('flags too few arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: 'Greet' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'b',
			edit: {
				span: {
					start: src.indexOf('"Ann"') + '"Ann"'.length,
					end: src.indexOf('"Ann"') + '"Ann"'.length,
				},
				newText: ', TODO_b',
			},
		});
	});

	it('checks argument counts for calls after numeric line labels', () => {
		const src =
			'Sub Main()\n' +
			'10 Greet "Ann"\n' +
			'20 Call Greet("Ann")\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-count', [
			{ span: 'Greet', message: 'expected 2 arguments' },
			{ span: 'Greet', message: 'expected 2 arguments' },
		]);
	});

	it('flags too many arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob", "Cat"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count');
	});

	it('flags a lone-identifier call that omits required arguments', () => {
		const src =
			'Sub Main()\n' +
			'    Greet\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count');
	});

	it('flags a runtime function call that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: 'MsgBox' });
		const argStart = src.indexOf('MsgBox(') + 'MsgBox('.length;
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Prompt',
			edit: {
				span: {
					start: argStart,
					end: argStart,
				},
				newText: 'TODO_Prompt',
			},
		});
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not arity-check runtime signatures when source names shadow them', () => {
		const src =
			'Private Const Format = 1\n' +
			'Sub Main()\n' +
			'    Dim MsgBox\n' +
			'    a = Format()\n' +
			'    b = MsgBox()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a parenless runtime function statement that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: 'MsgBox' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Prompt',
			edit: {
				span: {
					start: src.indexOf('MsgBox') + 'MsgBox'.length,
					end: src.indexOf('MsgBox') + 'MsgBox'.length,
				},
				newText: ' TODO_Prompt',
			},
		});
	});

	it('validates an explicit Call statement', () => {
		const src =
			'Sub Main()\n' +
			'    Call Greet("Ann", "Bob", "Cat")\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(1);
	});

	it('flags an omitted required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(total, )\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { severity: 'error', span: ',' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'TaxRate',
			edit: {
				span: {
					start: src.indexOf(', )'),
					end: src.indexOf(', )') + 2,
				},
				newText: ', TODO_TaxRate',
			},
		});
	});

	it('flags an omitted leading required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(, 0.08)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: ',' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Subtotal',
			edit: {
				span: {
					start: src.indexOf('(,') + 1,
					end: src.indexOf('(,') + 3,
				},
				newText: 'TODO_Subtotal, ',
			},
		});
	});

	it('accepts an omitted Optional argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, Optional ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(total, )\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('accepts a correct argument count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('honours Optional parameters', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'    Greet "Ann", "Bob"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, Optional ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('honours ParamArray (no upper bound)', () => {
		const src =
			'Sub Main()\n' +
			'    Log "a", "b", "c", "d"\n' +
			'End Sub\n' +
			'Sub Log(ParamArray items() As Variant)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('treats CallByName trailing args as a runtime ParamArray', () => {
		const src =
			'Sub Main()\n' +
			'    CallByName target, "Run", VbMethod, a0, a1, a2, a3\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('treats VBA Array ArgList as a runtime ParamArray', () => {
		const src =
			'Sub Main()\n' +
			'    Dim monthNames As Variant\n' +
			'    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _\n' +
			'                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", _\n' +
			'                       "Jan+1", "Feb+1", "Mar+1", "Apr+1", "May+1", _\n' +
			'                       "Jun+1", "Jul+1", "Aug+1")\n' +
			'    monthNames = Array()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('validates named-argument names but not the count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet who:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'who' });
	});

	it('flags duplicate named arguments', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs alpha:=1, alpha:=2\n' +
			'    NamedArgs alpha:=1, ALPHA:=2\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');

		expectDiagnostics(src, hits, 'argument-count', [{ span: 'alpha' }, { span: 'ALPHA' }]);
	});

	it('accepts a valid named argument', () => {
		const src =
			'Sub Main()\n' +
			'    Greet a:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, Optional ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check calls to unknown or cross-module procedures', () => {
		const src =
			'Sub Main()\n' +
			'    SomethingElse 1, 2, 3\n' +
			'    MsgBox "hi", vbOKOnly, "title", 0, 0\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check runtime statements whose curated signature has no parameter list', () => {
		const src = 'Sub Main()\n    Randomize 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check an ambiguous (duplicated) procedure name', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('uses a unique exported project signature for cross-module argument count', () => {
		const caller =
			'Public Sub Main()\n' +
			'    PrintTotal 100\n' +
			'End Sub\n';
		const helper =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helper },
			], 'Caller'),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'PrintTotal' });
	});

	it('validates same-module Declare argument counts', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Sleep' });
	});

	it('uses the active conditional Declare signature for same-module calls', () => {
		const src =
			'#If VBA7 Then\n' +
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'#Else\n' +
			'Private Declare Sub Sleep Lib "kernel32" ()\n' +
			'#End If\n' +
			'Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const vba7Hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: true } },
			}),
			'argument-count',
		);
		expect(vba7Hits).toHaveLength(1);
		expect(spanText(src, vba7Hits[0])).toBe('Sleep');

		const legacyHits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: false } },
			}),
			'argument-count',
		);
		expect(legacyHits).toHaveLength(0);
	});

	it('uses exported project Declare signatures for cross-module argument count', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const nativeApi =
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'NativeApi', source: nativeApi },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'Sleep' });
	});

	it('does not arity-check ambiguous exported project signatures', () => {
		const caller =
			'Public Sub Main()\n' +
			'    PrintTotal 100\n' +
			'End Sub\n';
		const first = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const second =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures: projectProcedures([
						{ moduleName: 'Caller', source: caller },
						{ moduleName: 'First', source: first },
						{ moduleName: 'Second', source: second },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('uses a module-qualified project signature even when the bare name is ambiguous', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.PrintTotal 100\n' +
			'End Sub\n';
		const helpers =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const alternate = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Helpers', source: helpers },
					{ moduleName: 'Alternate', source: alternate },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'PrintTotal' });
	});

	it('keeps module-qualified project diagnostics stable under module order changes', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.PrintTotal 100\n' +
			'End Sub\n';
		const helpers =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const alternate = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const ordered = projectProcedures([
			{ moduleName: 'Caller', source: caller },
			{ moduleName: 'Helpers', source: helpers },
			{ moduleName: 'Alternate', source: alternate },
		]);
		const reversed = projectProcedures([
			{ moduleName: 'Alternate', source: alternate },
			{ moduleName: 'Helpers', source: helpers },
			{ moduleName: 'Caller', source: caller },
		]);
		const messages = (projectProcedures: ReturnType<ProjectIndex['procedureSignatures']>) =>
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures,
				}),
				'argument-count',
			).map((hit) => `${spanText(caller, hit)}:${hit.message}`);
		expect(messages(ordered)).toEqual(messages(reversed));
	});

	it('does not validate a module-qualified private project procedure', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.Hidden 100\n' +
			'End Sub\n';
		const helpers =
			'Private Sub Hidden(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures: projectProcedures([
						{ moduleName: 'Caller', source: caller },
						{ moduleName: 'Helpers', source: helpers },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('flags missing required arguments on generated host member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Dim wb As Workbook\n' +
			'    Set wb = Workbooks.Open()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Open' });
	});

	it('flags extra arguments on generated host member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Call Application.Calculate(1)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Calculate' });
	});

	it('flags missing required arguments on host members with fallback signatures', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Range\n' +
			'    Set r = ActiveSheet.Range()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Range' });
	});

	it('does not treat collection indexing as member-call arity', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Range\n' +
			'    Set r = Workbooks(1).Sheets(1).Range()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Range' });
	});

	it('accepts correct host member argument counts', () => {
		const src =
			'Sub Main()\n' +
			'    Dim wb As Workbook\n' +
			'    Application.Calculate\n' +
			'    Call Application.Calculate()\n' +
			'    Set wb = Workbooks.Open("Book1.xlsx")\n' +
			'    ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not treat zero-argument property result indexing as member-call arity', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    Dim child As Object\n' +
			'    Set child = p.Children(1)\n' +
			'End Sub\n';
		const person =
			'Public Property Get Children() As Collection\n' +
			'End Property\n';
		expect(
			byCode(
				analyzeModule(caller, {
					projectClassMembers: projectClassMembers([
						{ moduleName: 'Person', moduleKind: 'class', source: person },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('accepts Excel collection default Item calls returned from properties', () => {
		const src =
			'Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Dim loTests As ListObject\n' +
			'    Dim lcPassed As ListColumn\n' +
			'    Set ws = ActiveSheet\n' +
			'    Set loTests = ws.ListObjects("Tests")\n' +
			'    Set lcPassed = loTests.ListColumns("Passed")\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);

		expect(byCode(diagnostics, 'argument-count')).toHaveLength(0);
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
		expect(diagnostics.filter((diag) => diag.severity === 'error').map((diag) => ({
			code: diag.code,
			text: spanText(src, diag),
			message: diag.message,
		}))).toEqual([]);
	});

	it('flags missing required arguments on source-backed class member calls', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    Call p.Save()\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'Save' });
	});

	it('honours ParamArray on source-backed class member calls', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim cb As Callback\n' +
			'    cb.Run first, second, third\n' +
			'End Sub\n';
		const callback =
			'Public Function Run(ParamArray params() As Variant) As Variant\n' +
			'End Function\n';

		expect(
			byCode(
				analyzeModule(caller, {
					projectClassMembers: projectClassMembers([
						{ moduleName: 'Callback', moduleKind: 'class', source: callback },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('validates parenless source-backed class member call statements', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.Save\n' +
			'    p.Save "ok"\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'Save' });
	});

	it('validates parenless generated host member call statements', () => {
		const src =
			'Sub Main()\n' +
			'    ActiveSheet.Range\n' +
			'    ActiveSheet.Range "A1"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Range' });
	});

	it('validates required arguments on runtime object member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Err.Raise\n' +
			'    Err.Raise vbObjectError + 1\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Raise' });
	});

	it('accepts Debug.Print output lists and validates Debug.Assert arity', () => {
		const src =
			'Sub Main()\n' +
			'    Debug.Print "value", 1, True\n' +
			'    Debug.Assert\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Assert' });
	});

	it('flags missing required arguments on current class Me member calls', () => {
		const src =
			'Public Sub Main()\n' +
			'    Call Me.Save()\n' +
			'End Sub\n' +
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Person',
				moduleKind: 'class',
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: src },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(src, hits, 'argument-count', { span: 'Save' });
	});
});

describe('analyzeModule - argument type validation', () => {
	it('flags a nonnumeric string literal passed to a same-module numeric parameter', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', {
			severity: 'error',
			span: '"blah"',
		});
	});

	it('flags intrinsic CVErr Error Variants passed to scalar parameters', () => {
		const src =
			'Public Sub NeedsLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub NeedsVariant(ByVal value As Variant)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsLong CVErr(2015)\n' +
			'    NeedsLong VBA.CVErr(2015)\n' +
			'    NeedsVariant CVErr(2015)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');

		expectDiagnostics(src, hits, 'argument-type-mismatch', [
			{ span: 'CVErr(2015)' },
			{ span: 'VBA.CVErr(2015)' },
		]);
	});

	it('flags Null passed to scalar parameters', () => {
		const src =
			'Public Sub NeedsLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub NeedsString(ByVal value As String)\n' +
			'End Sub\n' +
			'Public Sub NeedsVariant(ByVal value As Variant)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsLong Null\n' +
			'    NeedsString Null\n' +
			'    NeedsVariant Null\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');

		expectDiagnostics(src, hits, 'argument-type-mismatch', [{ span: 'Null' }, { span: 'Null' }]);
	});

	it('lets a source function named CVErr shadow the intrinsic in argument expressions', () => {
		const src =
			'Private Function CVErr(ByVal errorNumber As Long) As Long\n' +
			'    CVErr = errorNumber\n' +
			'End Function\n' +
			'Public Sub NeedsLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsLong CVErr(2015)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('uses a unique exported project signature for cross-module argument types', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'argument-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-type-mismatch', { span: '"blah"' });
	});

	it('validates same-module Declare argument types', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Public Sub T()\n' +
			'    Sleep "bad"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"bad"' });
	});

	it('uses module-qualified project signatures for argument types', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = Invoices.InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Invoices', source: invoices },
				]),
			}),
			'argument-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-type-mismatch', { span: '"blah"' });
	});

	it('uses a unique exported project function return type in nested calls', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel()\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('uses visible exported scalar globals as bare argument value types', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject SharedText\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'argument-object-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'SharedText' });
	});

	it('uses module-qualified exported scalar globals as argument value types', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Globals.SharedText\n' +
			'    NeedsObject Globals.SharedText & ""\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'argument-object-type-mismatch',
		);

		expectDiagnostics(caller, hits, 'argument-object-type-mismatch', [
			{ span: 'SharedText', message: 'Globals.SharedText As String' },
			{ span: 'Globals.SharedText & ""', message: 'string concatenation expression' },
		]);
	});

	it('keeps unknown module-qualified value types quiet for argument inference', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Missing.SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(caller, [], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('keeps local shadows and ambiguous exported globals quiet for bare argument value types', () => {
		const shadowCaller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim SharedText\n' +
			'    NeedsObject SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);

		const ambiguousCaller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses a module-qualified project function return type in nested calls', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Labels.MakeLabel()\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('uses source-backed member return types in argument expressions', () => {
		const person = 'Public Property Get Name() As String\nEnd Property\n';
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    NeedsObject p.Name\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'argument-object-type-mismatch', { span: 'Name' });
	});

	it('does not infer bare member functions that require arguments as values', () => {
		const person =
			'Public Function Name(ByVal index As Long) As String\n' +
			'End Function\n';
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    NeedsObject p.Name\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses member return types inside string concatenation argument expressions', () => {
		const person = 'Public Property Get Name() As String\nEnd Property\n';
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    NeedsObject p.Name & "!"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'argument-object-type-mismatch', { span: 'p.Name & "!"' });
	});

	it('accepts numeric literals and numeric string literals for numeric parameters', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    a = InvoiceTotal(100, 0.08)\n' +
			'    b = InvoiceTotal("100", 0.08)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('errors on decimal literals outside Byte and Integer parameter bounds', () => {
		const src =
			'Public Sub TakesByte(ByVal value As Byte)\n' +
			'End Sub\n' +
			'Public Sub TakesInteger(ByVal value As Integer)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    TakesByte 0\n' +
			'    TakesByte 255\n' +
			'    TakesByte 256\n' +
			'    TakesByte -1\n' +
			'    TakesInteger -32768\n' +
			'    TakesInteger 32767\n' +
			'    TakesInteger 32768\n' +
			'    TakesInteger -32769\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-type-mismatch', [
			{ span: '256', message: ['Byte', "Run-time error '6'"] },
			{ span: '-1', message: ['Byte', "Run-time error '6'"] },
			{ span: '32768', message: ['Integer', "Run-time error '6'"] },
			{ span: '-32769', message: ['Integer', "Run-time error '6'"] },
		]);
	});

	it('does not warn on string variables whose runtime value is unknown', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim label As String\n' +
			'    total = InvoiceTotal(label)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('flags known scalar variable type mismatches for ByRef parameters', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'byref-argument-type-mismatch', { span: 'amount' });
	});

	it('treats omitted parameter passing markers as default ByRef', () => {
		const src =
			'Public Sub Mutate(value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(1);
	});

	it('flags known Object variable mismatches for ByRef scalar parameters', () => {
		const src =
			'Public Sub NeedsString(ByRef value As String)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    NeedsString obj\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'byref-argument-type-mismatch');

		expectDiagnostic(src, hits, 'byref-argument-type-mismatch', { span: 'obj' });
	});

	it('does not apply ByRef exactness to ByVal parameters, literals, or expression temporaries', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub ReadValue(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub ReadString(ByVal value As String)\n' +
			'End Sub\n' +
			'Public Sub NeedsString(ByRef value As String)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Dim longAmount As Long\n' +
			'    Dim flexible As Variant\n' +
			'    Dim obj As Object\n' +
			'    ReadValue amount\n' +
			'    ReadString obj\n' +
			'    NeedsString flexible\n' +
			'    Mutate 1\n' +
			'    Mutate (longAmount)\n' +
			'    Mutate longAmount + 1\n' +
			'    Mutate Range("A1").Value\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(0);
	});

	it('uses module #Const activity when checking ByRef exactness in conditional branches', () => {
		const src =
			'#Const NegActiveBranch = True\n' +
			'#Const NegInactiveBranch = False\n' +
			'Public Sub NegSupport01_NeedsLong(ByRef value As Long)\n' +
			'End Sub\n' +
			'#If NegActiveBranch Then\n' +
			'Public Sub NegCond01_ActiveInvalidBranch()\n' +
			'    Dim amount As Integer\n' +
			'    NegSupport01_NeedsLong amount\n' +
			'End Sub\n' +
			'#Else\n' +
			'Public Sub NegCond02_InactiveAlternateBranch()\n' +
			'End Sub\n' +
			'#End If\n' +
			'#If NegInactiveBranch Then\n' +
			'Public Sub NegCond03_InactiveInvalidBranch()\n' +
			'    Dim amount As Integer\n' +
			'    NegSupport01_NeedsLong amount\n' +
			'End Sub\n' +
			'#Else\n' +
			'Public Sub NegCond04_ActiveValidBranch()\n' +
			'End Sub\n' +
			'#End If\n';
		const hits = byCode(analyzeModule(src), 'byref-argument-type-mismatch');

		expectDiagnostic(src, hits, 'byref-argument-type-mismatch', { span: 'amount' });
	});

	it('uses unique exported project signatures for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		const helpers = 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helpers },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('amount');
	});

	it('uses visible exported project variables for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Mutate SharedAmount\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
				{ moduleName: 'Globals', source: 'Public SharedAmount As Integer\n' },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'byref-argument-type-mismatch', { span: 'SharedAmount' });
	});

	it('uses module-qualified exported project variables for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Mutate Globals.SharedAmount\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
				{ moduleName: 'Globals', source: 'Public SharedAmount As Integer\n' },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'byref-argument-type-mismatch', {
			span: 'Globals.SharedAmount',
		});
	});

	it('keeps local shadows and ambiguous exported variables quiet for ByRef exactness', () => {
		const shadowCaller =
			'Public Sub T()\n' +
			'    Dim SharedAmount As Variant\n' +
			'    Mutate SharedAmount\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
			{ moduleName: 'Globals', source: 'Public SharedAmount As Integer\n' },
		], 'Caller'), 'byref-argument-type-mismatch')).toHaveLength(0);

		const ambiguousCaller =
			'Public Sub T()\n' +
			'    Mutate SharedAmount\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
			{ moduleName: 'GlobalsA', source: 'Public SharedAmount As Integer\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedAmount As String\n' },
		], 'Caller'), 'byref-argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn on Variant arguments whose runtime value is unknown', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim value As Variant\n' +
			'    total = InvoiceTotal(value)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('maps named arguments to the named parameter type', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal(TaxRate:=0.08, Subtotal:="blah")\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"blah"' });
	});

	it('validates selected native VBA runtime parameter types', () => {
		const src = 'Sub T()\n    x = Left("abcdef", "bad")\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"bad"' });
	});

	it('does not validate runtime parameter types when a local shadows the runtime function', () => {
		const src =
			'Sub T()\n' +
			'    Dim Left As Variant\n' +
			'    value = Left("abcdef", "bad")\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'runtime-argument-value')).toHaveLength(0);
	});

	it('flags negative literal values for selected native VBA runtime argument bounds', () => {
		const src =
			'Sub T()\n' +
			'    a = Left$("abcdef", -1)\n' +
			'    b = Left("abcdef", -2)\n' +
			'    c = String$(-3, "x")\n' +
			'    d = String(-4, "x")\n' +
			'    e = VBA.Left$("abcdef", -5)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '-1', message: ['Left$', 'Length', "Run-time error '5'"] },
			{ span: '-2' },
			{ span: '-3', message: ['String$', 'Number'] },
			{ span: '-4' },
			{ span: '-5' },
		]);
	});

	it('flags oracle-backed runtime argument bounds for Right Space and Mid', () => {
		const src =
			'Sub T()\n' +
			'    a = Right$("abcdef", -1)\n' +
			'    b = Right("abcdef", -2)\n' +
			'    c = Space$(-3)\n' +
			'    d = Space(-4)\n' +
			'    e = Mid$("abcdef", 0, 1)\n' +
			'    f = Mid("abcdef", -1, 1)\n' +
			'    g = Mid$("abcdef", 1, -5)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '-1', message: ['Right$', 'Length'] },
			{ span: '-2' },
			{ span: '-3', message: ['Space$', 'Number'] },
			{ span: '-4' },
			{ span: '0', message: ['Mid$', 'Start'] },
			{ span: '-1' },
			{ span: '-5', message: 'Length' },
		]);
	});

	it('flags oracle-backed runtime argument bounds for Replace', () => {
		const src =
			'Sub T()\n' +
			'    a = Replace("abcdef", "a", "z", 0)\n' +
			'    b = Replace("abcdef", "a", "z", -1)\n' +
			'    c = Replace("aaaa", "a", "z", 1, -2)\n' +
			'    d = Replace("aaaa", "a", "z", Count:=-2, Start:=1)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '0', message: ['Replace', 'Start'] },
			{ span: '-1' },
			{ span: '-2', message: 'Count' },
			{ span: '-2' },
		]);
	});

	it('folds deterministic Const Enum and integer expressions for runtime argument bounds', () => {
		const src =
			'Private Const BadLength As Long = -1\n' +
			'Private Const GoodLength As Long = 0\n' +
			'Private Enum RuntimeArgStart\n' +
			'    EnumBadStart = 0\n' +
			'End Enum\n' +
			'Sub T()\n' +
			'    Const BadStart As Long = 1 - 1\n' +
			'    Const BadCount As Long = -1 - 1\n' +
			'    Const GoodCount As Long = 0 - 1\n' +
			'    a = Left$("abcdef", BadLength)\n' +
			'    b = Left$("abcdef", GoodLength)\n' +
			'    c = Left$("abcdef", 1 - 2)\n' +
			'    d = Right$("abcdef", 1 - 1)\n' +
			'    e = Replace("abcdef", "a", "z", BadStart)\n' +
			'    f = Replace("aaaa", "a", "z", 1, BadCount)\n' +
			'    g = Replace("aaaa", "a", "z", 1, GoodCount)\n' +
			'    h = Mid$("abcdef", EnumBadStart, 1)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: 'BadLength', message: 'is -1' },
			{ span: '1 - 2', message: 'is -1' },
			{ span: 'BadStart', message: 'is 0' },
			{ span: 'BadCount', message: 'is -2' },
			{ span: 'EnumBadStart', message: 'is 0' },
		]);
	});

	it('folds verified runtime and host constants for runtime argument bounds', () => {
		const src =
			'Sub T()\n' +
			'    a = Left$("abcdef", vbFalse - 1)\n' +
			'    b = Left$("abcdef", VBA.vbFalse - 1)\n' +
			'    c = Left$("abcdef", xlAbove - 1)\n' +
			'    d = Left$("abcdef", xlNo - 1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');

		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'vbFalse - 1',
			'VBA.vbFalse - 1',
			'xlAbove - 1',
		]);
	});

	it('lets source declarations shadow verified runtime and host constants for runtime argument bounds', () => {
		const src =
			'Private Const vbFalse As Long = 1\n' +
			'Sub T()\n' +
			'    Dim xlAbove As Long\n' +
			'    xlAbove = 1\n' +
			'    a = Left$("abcdef", vbFalse - 1)\n' +
			'    b = Left$("abcdef", xlAbove - 1)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'runtime-argument-value')).toHaveLength(0);
	});

	it('flags oracle-backed InStr Start values below one without flagging two-argument calls', () => {
		const src =
			'Sub T()\n' +
			'    Const BadStart As Long = 0\n' +
			'    a = InStr(0, "abcdef", "a")\n' +
			'    b = InStr(-1, "abcdef", "a")\n' +
			'    c = InStr(1, "abcdef", "a")\n' +
			'    d = InStr("abcdef", "a")\n' +
			'    e = InStr(BadStart, "abcdef", "a")\n' +
			'    f = InStr(Start:=0, String1:="abcdef", String2:="a")\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '0', message: "'Start' of 'InStr' is 0" },
			{ span: '-1', message: "'Start' of 'InStr' is -1" },
			{ span: 'BadStart', message: "'Start' of 'InStr' is 0" },
		]);
	});

	it('flags oracle-backed Chr and ChrW CharCode values outside proven bounds', () => {
		const src =
			'Sub T()\n' +
			'    Const BadChrLow As Long = -1\n' +
			'    Const BadChrHigh As Long = 256\n' +
			'    Const BadChrWHigh As Long = 65536\n' +
			'    a = Chr(-1)\n' +
			'    b = Chr(0)\n' +
			'    c = Chr(255)\n' +
			'    d = Chr(256)\n' +
			'    e = Chr(BadChrLow)\n' +
			'    f = Chr(BadChrHigh)\n' +
			'    g = ChrW(-1)\n' +
			'    h = ChrW(65535)\n' +
			'    i = ChrW(65536)\n' +
			'    j = ChrW(BadChrWHigh)\n' +
			'    k = VBA.Chr(256)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '-1', message: "'CharCode' of 'Chr' is -1" },
			{ span: '256', message: "'CharCode' of 'Chr' is 256" },
			{ span: 'BadChrLow', message: "'CharCode' of 'Chr' is -1" },
			{ span: 'BadChrHigh', message: "'CharCode' of 'Chr' is 256" },
			{ span: '65536', message: "'CharCode' of 'ChrW' is 65536" },
			{ span: 'BadChrWHigh', message: "'CharCode' of 'ChrW' is 65536" },
			{ span: '256', message: "'CharCode' of 'Chr' is 256" },
		]);
	});

	it('folds visible cross-module Const and Enum values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Left$("abcdef", SharedGoodLength)\n' +
			'    c = Mid$("abcdef", SharedBadStart, 1)\n' +
			'    d = Replace("aaaa", "a", "z", 1, SharedBadCount)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Const SharedGoodLength As Long = 0\n' +
						'Public Const SharedBadCount As Long = -2\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedBadLength',
			'SharedBadStart',
			'SharedBadCount',
		]);
	});

	it('folds module-qualified cross-module Const and Enum values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedRuntimeArgs.SharedBadLength)\n' +
			'    b = Left$("abcdef", SharedRuntimeArgs.SharedGoodLength)\n' +
			'    c = Mid$("abcdef", SharedRuntimeArgs.SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Const SharedGoodLength As Long = 0\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedRuntimeArgs.SharedBadLength',
			'SharedRuntimeArgs.SharedBadStart',
		]);
	});

	it('folds exported cross-module Const values through private same-module helpers for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedBadLengthFromHidden)\n' +
			'    b = Left$("abcdef", SharedGoodLengthFromHidden)\n' +
			'    c = Left$("abcdef", SharedRuntimeArgs.SharedBadLengthFromHidden)\n' +
			'    d = Left$("abcdef", SharedRuntimeArgs.SharedGoodLengthFromHidden)\n' +
			'    e = Left$("abcdef", SharedRuntimeArgs.HiddenBadLength)\n' +
			'End Sub\n';
		const shared =
			'Private Const HiddenBadLength As Long = -1\n' +
			'Private Const HiddenGoodLength As Long = 0\n' +
			'Public Const SharedBadLengthFromHidden As Long = HiddenBadLength\n' +
			'Public Const SharedGoodLengthFromHidden As Long = HiddenGoodLength\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'SharedRuntimeArgs', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'runtime-argument-value',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedBadLengthFromHidden',
			'SharedRuntimeArgs.SharedBadLengthFromHidden',
		]);
	});

	it('defers hidden and ambiguous cross-module Const values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", HiddenBadLength)\n' +
			'    b = Left$("abcdef", AmbiguousLength)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'PrivateGlobals',
					source:
						'Private Const HiddenBadLength As Long = -1\n' +
						'Public Const AmbiguousLength As Long = -1\n',
				},
				{
					moduleName: 'MoreGlobals',
					source: 'Public Const AmbiguousLength As Long = 0\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(0);
	});

	it('lets current-module and local Const values shadow cross-module runtime argument bounds', () => {
		const caller =
			'Private Const SharedBadLength As Long = 0\n' +
			'Sub T()\n' +
			'    Const SharedBadStart As Long = 1\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Mid$("abcdef", SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(0);
	});

	it('lets source variables shadow cross-module runtime argument constants', () => {
		const caller =
			'Private SharedBadLength As Long\n' +
			'Sub T()\n' +
			'    Dim SharedBadStart As Long\n' +
			'    SharedBadLength = 1\n' +
			'    SharedBadStart = 1\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Mid$("abcdef", SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);

		expect(hits).toHaveLength(0);
	});

	it('accepts zero and unknown runtime argument values for selected native bounds', () => {
		const src =
			'Sub T()\n' +
			'    Dim count As Long\n' +
			'    a = Left$("abcdef", 0)\n' +
			'    b = Left("abcdef", count)\n' +
			'    c = String$(0, "x")\n' +
			'    d = String(count, "x")\n' +
			'    e = Right$("abcdef", 0)\n' +
			'    f = Right("abcdef", count)\n' +
			'    g = Space$(0)\n' +
			'    h = Space(count)\n' +
			'    i = Mid$("abcdef", 1, 0)\n' +
			'    j = Mid("abcdef", count, count)\n' +
			'    k = Replace("abcdef", "a", "z", 1)\n' +
			'    l = Replace("aaaa", "a", "z", 1, -1)\n' +
			'    m = Replace("aaaa", "a", "z", 1, 0)\n' +
			'    n = Replace("aaaa", "a", "z", count, count)\n' +
			'    o = Replace$("aaaa", "a", "z", 0)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'runtime-argument-value')).toHaveLength(0);
	});

	it('does not treat shadowed Left calls as native runtime argument-value checks', () => {
		const src =
			'Public Function Left(ByVal text As String, ByVal count As Long) As String\n' +
			'End Function\n' +
			'Sub T()\n' +
			'    Dim localValue As Long\n' +
			'    Dim Left As Long\n' +
			'    a = Left("abcdef", -1)\n' +
			'    b = VBA.Left$("abcdef", -1)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: '-1' });
	});

	it('flags plainly invalid literal CDate conversions', () => {
		const src =
			'Sub T()\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-conversion-value');

		expectDiagnostics(src, hits, 'runtime-conversion-value', [
			{ span: '"not a date"', message: ['CDate', "Run-time error '13'"] },
			{ span: '""', message: 'VBA.CDate' },
		]);
	});

	it('keeps date-looking, localized, variable, and non-CDate conversions quiet', () => {
		const src =
			'Sub T()\n' +
			'    Dim Value As Date\n' +
			'    Dim text As String\n' +
			'    Value = CDate("1/2/2020")\n' +
			'    Value = CDate("March")\n' +
			'    Value = CDate("Mar 1")\n' +
			'    Value = CDate("März")\n' +
			'    Value = CDate(text)\n' +
			'    Value = DateValue("not a date")\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'runtime-conversion-value')).toHaveLength(0);
	});

	it('does not treat shadowed CDate calls as native conversion checks', () => {
		const src =
			'Sub T()\n' +
			'    Dim CDate As Variant\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("not a date")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-conversion-value');

		expectDiagnostic(src, hits, 'runtime-conversion-value', { span: '"not a date"' });
	});

	it('does not treat project-visible source names as native conversion checks', () => {
		const globals = 'Public Const CDate As Long = 1\n';
		const caller =
			'Sub T()\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("not a date")\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(
			caller,
			[
				{ moduleName: 'Globals', moduleKind: 'standard', source: globals },
				{ moduleName: 'Caller', moduleKind: 'standard', source: caller },
			],
			'Caller',
		);
		const hits = byCode(diagnostics, 'runtime-conversion-value');

		expectDiagnostic(caller, hits, 'runtime-conversion-value', { span: '"not a date"' });
	});

	it('keeps Close of a literal file number quiet', () => {
		const src = 'Option Explicit\nSub T()\n    Close #99\nEnd Sub\n';
		expect(analyzeModule(src)).toEqual([]);
	});

	it('keeps string-suffixed Left and Right intrinsics quiet', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Debug.Print Left$("abcdef", 2)\n' +
			'    Debug.Print Left("abcdef", 2)\n' +
			'    Debug.Print Right$("abcdef", 2)\n' +
			'    Debug.Print Right("abcdef", 2)\n' +
			'End Sub\n';
		expect(analyzeModule(src)).toEqual([]);
	});

	it('does not infer runtime parameter types from display names', () => {
		const src = 'Sub T()\n    Randomize "bad"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn when a type is unknown or Variant-like', () => {
		const src = 'Sub T()\n    x = Format("blah", "0.00")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('uses same-module function return types as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			severity: 'error',
			span: 'MakeLabel',
		});
	});

	it('uses string-suffixed runtime aliases as argument types', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Left$("abcdef", 2)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', { span: 'Left$' });
	});

	it('uses verified runtime and host constants as argument value types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject vbFalse\n' +
			'    NeedsObject VBA.vbFalse\n' +
			'    NeedsObject xlAbove\n' +
			'    NeedsObject vbNullString\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');

		expectDiagnostics(src, hits, 'argument-object-type-mismatch', [
			{ span: 'vbFalse', message: 'vbFalse As VbTriState' },
			{ span: 'vbFalse', message: 'VBA.vbFalse As VbTriState' },
			{ span: 'xlAbove', message: 'xlAbove As Constants' },
			{ span: 'vbNullString', message: 'vbNullString As String' },
		]);
	});

	it('lets source declarations shadow verified external constants for argument value types', () => {
		const src =
			'Public Function vbFalse() As Object\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim xlAbove As Object\n' +
			'    NeedsObject vbFalse\n' +
			'    NeedsObject xlAbove\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses parameterless Function and Property Get references as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Property Get CurrentLabel() As String\n' +
			'End Property\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel\n' +
			'    NeedsObject (CurrentLabel)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-object-type-mismatch', [
			{ span: 'MakeLabel', message: 'MakeLabel As String' },
			{ span: 'CurrentLabel', message: 'CurrentLabel As String' },
		]);
	});

	it('keeps the current Function return variable available for bare argument inference', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Function CurrentLabel(ByVal index As Long) As String\n' +
			'    NeedsObject CurrentLabel\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');

		expectDiagnostic(src, hits, 'argument-object-type-mismatch', { span: 'CurrentLabel' });
	});

	it('uses module-qualified parameterless project Function references as argument types', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Labels.MakeLabel\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('does not infer ambiguous parameterless project Function references', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel\n' +
			'End Sub\n';
		const first = 'Public Function MakeLabel() As String\nEnd Function\n';
		const second = 'Public Function MakeLabel() As String\nEnd Function\n';
		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'FirstLabels', source: first },
			{ moduleName: 'SecondLabels', source: second },
		], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses generated host member signatures for scalar argument types', () => {
		const src =
			'Sub T()\n' +
			'    Call Application.DeleteCustomList("bad")\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"bad"' });
	});

	it('uses source-backed class member signatures for argument types', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.NeedsObject("bad")\n' +
			'End Sub\n';
		const person =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: '"bad"' });
	});

	it('uses parenless source-backed class member signatures for argument types', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.Save "bad"\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Count As Long)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-type-mismatch', { span: '"bad"' });
	});

	it('uses nested same-module function return types as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Function EchoLabel(ByVal value As String) As String\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject EchoLabel(MakeLabel())\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			span: 'EchoLabel',
		});
	});

	it('uses curated runtime conversion function return types as argument types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject CStr(123)\n' +
			'    NeedsObject CDbl("1")\n' +
			'    NeedsObject CCur(1)\n' +
			'    NeedsObject CLng(1)\n' +
			'    NeedsObject CBool("True")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(5);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'CStr',
			'CDbl',
			'CCur',
			'CLng',
			'CBool',
		]);
	});

	it('does not infer runtime return types when a local shadows the runtime function', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim Format\n' +
			'    NeedsObject Format(123)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses obvious numeric arithmetic expression return types as argument types', () => {
		const src =
			'Public Function TaxRate() As Double\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim subtotal As Currency\n' +
			'    Dim fee As Double\n' +
			'    NeedsObject subtotal + fee * TaxRate()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			span: 'subtotal + fee * TaxRate()',
		});
	});

	it('does not infer arithmetic expressions with unknown, Variant, or string operands', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Double\n' +
			'    Dim flexible As Variant\n' +
			'    Dim label As String\n' +
			'    NeedsObject amount + flexible\n' +
			'    NeedsObject amount + UnknownValue\n' +
			'    NeedsObject amount + label\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses string concatenation expression return types as argument types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim prefix As String\n' +
			'    Dim amount As Double\n' +
			'    NeedsObject prefix & amount & CStr(123)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			span: 'prefix & amount & CStr(123)',
		});
	});

	it('does not infer string concatenation expressions with unknown or Variant operands', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim prefix As String\n' +
			'    Dim flexible As Variant\n' +
			'    NeedsObject prefix & flexible\n' +
			'    NeedsObject prefix & UnknownValue\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - assignment type validation', () => {
	it('errors on a nonnumeric string literal assigned to a numeric variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim total As Double\n' +
			'    total = "blah"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-type-mismatch', {
			severity: 'error',
			span: '"blah"',
		});
	});

	it('errors on intrinsic CVErr Error Variants assigned to scalar variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim first As Long\n' +
			'    Dim second As Long\n' +
			'    Dim flexible As Variant\n' +
			'    first = CVErr(2015)\n' +
			'    second = VBA.CVErr(2015)\n' +
			'    flexible = CVErr(2015)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-type-mismatch', [
			{ span: 'CVErr(2015)' },
			{ span: 'VBA.CVErr(2015)' },
		]);
	});

	it('errors on Null assigned to scalar variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim number As Long\n' +
			'    Dim message As String\n' +
			'    Dim flexible As Variant\n' +
			'    number = Null\n' +
			'    message = Null\n' +
			'    flexible = Null\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-type-mismatch', [{ span: 'Null' }, { span: 'Null' }]);
	});

	it('lets a source function named CVErr shadow the intrinsic in assignment expressions', () => {
		const src =
			'Private Function CVErr(ByVal errorNumber As Long) As Long\n' +
			'    CVErr = errorNumber\n' +
			'End Function\n' +
			'Public Sub T()\n' +
			'    Dim value As Long\n' +
			'    value = CVErr(2015)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('errors on a non-Boolean string literal assigned to a Boolean variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim enabled As Boolean\n' +
			'    enabled = "maybe"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-type-mismatch', { span: '"maybe"' });
	});

	it('errors on decimal literals outside Byte and Integer assignment bounds', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim small As Byte\n' +
			'    Dim count As Integer\n' +
			'    small = 0\n' +
			'    small = 255\n' +
			'    small = 256\n' +
			'    small = -1\n' +
			'    count = -32768\n' +
			'    count = 32767\n' +
			'    count = 32768\n' +
			'    count = -32769\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'assignment-type-mismatch', [
			{ span: '256', message: ['Byte', "Run-time error '6'"] },
			{ span: '-1', message: ['Byte', "Run-time error '6'"] },
			{ span: '32768', message: ['Integer', "Run-time error '6'"] },
			{ span: '-32769', message: ['Integer', "Run-time error '6'"] },
		]);
	});

	it('accepts VBA scalar coercions and unknown assignment values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim total As Double\n' +
			'    Dim label As String\n' +
			'    Dim flexible As Variant\n' +
			'    total = "100"\n' +
			'    label = 123\n' +
			'    total = flexible\n' +
			'    total = UnknownValue\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('uses visible exported scalar globals for assignment target types', () => {
		const caller =
			'Public Sub T()\n' +
			'    SharedCount = "bad"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedCount As Long\n' },
			], 'Caller'),
			'assignment-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'assignment-type-mismatch', { span: '"bad"' });
	});

	it('does not leak a broader typed declaration through an untyped local assignment target', () => {
		const src =
			'Private Value As Long\n' +
			'Public Sub T()\n' +
			'    Dim Value\n' +
			'    Value = "not numeric"\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('flags array variables assigned to scalar variables', () => {
		const src =
			'Private ModuleValues(1 To 3) As Long\n' +
			'Public Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    Dim DynamicValues() As String\n' +
			'    Dim Value As Long\n' +
			'    Dim ScalarText As String\n' +
			'    Value = Values\n' +
			'    ScalarText = DynamicValues\n' +
			'    Value = ModuleValues\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'array-assignment-to-scalar');

		expectDiagnostics(src, hits, 'array-assignment-to-scalar', [
			{ severity: 'error', span: 'Values' },
			{ span: 'DynamicValues' },
			{ span: 'ModuleValues' },
		]);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('uses visible exported array globals for array assignment to scalar', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'    Value = SharedValues\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
			], 'Caller'),
			'array-assignment-to-scalar',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedValues');
	});

	it('does not leak exported array shapes through local untyped assignment source shadows', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim SharedValues\n' +
			'    Dim Value As Long\n' +
			'    Value = SharedValues\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported array assignment sources quiet', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'    Value = SharedValues\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValues() As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('does not treat visible exported array assignment targets as scalar targets', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    SharedValues = Values\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('accepts array assignments to Variant, array targets, indexed elements, and unknown values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    Dim Target() As Long\n' +
			'    Dim Flexible As Variant\n' +
			'    Dim Value As Long\n' +
			'    Flexible = Values\n' +
			'    Target = Values\n' +
			'    Value = Values(1)\n' +
			'    Value = UnknownValues\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('accepts array assignments to array-returning Function and Property Get return variables', () => {
		const src =
			'Public Function Names() As String()\n' +
			'    Dim values() As String\n' +
			'    Names = values\n' +
			'End Function\n' +
			'\n' +
			'Public Property Get MeasuresKeys() As String()\n' +
			'    Dim sOut() As String\n' +
			'    MeasuresKeys = sOut\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking array assignment to scalar', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'#If VBA7 Then\n' +
			'    Dim Values() As Long\n' +
			'#Else\n' +
			'    Dim Values As Long\n' +
			'#End If\n' +
			'    Value = Values\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'array-assignment-to-scalar',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'array-assignment-to-scalar',
			),
		).toHaveLength(0);
	});

	it('flags scalar variables passed to array bound functions', () => {
		const src =
			'Private ModuleValue As String\n' +
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'    Debug.Print LBound(Value)\n' +
			'    Debug.Print UBound(ModuleValue, 1)\n' +
			'    Debug.Print VBA.LBound(Value)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'array-bound-requires-array');

		expectDiagnostics(src, hits, 'array-bound-requires-array', [
			{ severity: 'error', span: 'Value' },
			{ span: 'ModuleValue' },
			{ span: 'Value' },
		]);
	});

	it('uses visible exported scalar globals for array bound argument shapes', () => {
		const caller =
			'Public Sub T()\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
			], 'Caller'),
			'array-bound-requires-array',
		);

		expectDiagnostic(caller, hits, 'array-bound-requires-array', { span: 'SharedValue' });
	});

	it('does not leak exported scalar shapes through local untyped array-bound shadows', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim SharedValue\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-bound-requires-array')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported scalar array-bound arguments quiet', () => {
		const caller =
			'Public Sub T()\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-bound-requires-array')).toHaveLength(0);
	});

	it('accepts array, Variant, unknown, and member-expression array bound arguments', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    Dim FixedValues(1 To 3) As Long\n' +
			'    Dim Flexible As Variant\n' +
			'    Debug.Print LBound(Values)\n' +
			'    Debug.Print UBound(FixedValues, 1)\n' +
			'    Debug.Print LBound(Flexible)\n' +
			'    Debug.Print UBound(UnknownValues)\n' +
			'    Debug.Print LBound(Settings.Values)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'array-bound-requires-array')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking array bound functions', () => {
		const src =
			'Public Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    Debug.Print LBound(Values)\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'array-bound-requires-array',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'array-bound-requires-array',
			),
		).toHaveLength(0);
	});

	it('requires Set when assigning to a known object variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws = ActiveSheet\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'set-required', { span: 'ws' });
	});

	it('reports missing Set for Object variables without treating it as scalar coercion', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    item = "blah"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'set-required')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('checks scalar Function return assignment types', () => {
		const src =
			'Public Function Total() As Double\n' +
			'    Total = "blah"\n' +
			'End Function\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-type-mismatch', { span: '"blah"' });
	});

	it('requires Set for object Function return assignments', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    MakePerson = New Person\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MakePerson');
	});

	it('checks object Function return assignment compatibility', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Class1\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Class1' });
	});

	it('accepts compatible object Function return assignment', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Person\n' +
			'End Function\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: '' },
			]),
		});
		expect(byCode(diagnostics, 'set-required')).toHaveLength(0);
		expect(byCode(diagnostics, 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('checks Property Get return assignments like Function returns', () => {
		const src =
			'Public Property Get Child() As Person\n' +
			'    Child = New Person\n' +
			'End Property\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
	});

	it('flags Set assignment between incompatible project class types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p = New Class1\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Class1' });
	});

	it('accepts Set assignment to a project interface implemented by another class', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Class1\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: 'Public Sub Save()\nEnd Sub\n' },
				{ moduleName: 'Class1', moduleKind: 'class', source: 'Implements Person\n' },
			]),
		});
		expect(byCode(diagnostics, 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('flags scalar Set assignment to a project object variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = "bad"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: '"bad"' });
	});

	it('flags scalar Set assignment from a parameterless Function reference', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = MakeLabel\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('uses source-backed member return types in Set assignment expressions', () => {
		const person = 'Public Property Get Name() As String\nEnd Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Dim child As Person\n' +
			'    Set child = p.Name\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Name' });
	});

	it('uses host member-call return types in Set assignment compatibility', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim wb As Workbook\n' +
			'    Dim rng As Range\n' +
			'    Set rng = ActiveSheet.Range("A1")\n' +
			'    Set wb = ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-object-type-mismatch', { span: 'Range' });
	});

	it('errors on a nonnumeric string literal assigned to a typed class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "blah"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-type-mismatch', { span: '"blah"' });
	});

	it('accepts numeric string assignment to a typed class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "2"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'assignment-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(0);
	});

	it('checks assignment types for public class fields', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "blah"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-type-mismatch', { span: '"blah"' });
	});

	it('accepts compatible assignments to public class fields', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "2"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'assignment-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(0);
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('uses visible UDT fields for missing-member and assignment type diagnostics', () => {
		const types = 'Public Type TPoint\n    X As Long\nEnd Type\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As TPoint\n' +
			'    p.X = "bad"\n' +
			'    p.Missing = 1\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			moduleName: 'Caller',
			projectClassMembers: projectMemberSurfaces(
				[
					{ moduleName: 'Caller', source: src },
					{ moduleName: 'Types', source: types },
				],
				'Caller',
			),
		});
		expectDiagnostic(src, diagnostics, 'assignment-type-mismatch', { span: '"bad"' });
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'Missing' });
	});

	it('requires Set for object-valued public class fields', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-required',
		);
		expectDiagnostic(src, hits, 'set-required', { span: 'Child' });
	});

	it('accepts Set for object-valued public class fields', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Person\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'set-required')).toHaveLength(0);
		expect(byCode(diagnostics, 'set-requires-object')).toHaveLength(0);
	});

	it('requires Set for Property Set object members', () => {
		const person =
			'Private mChild As Person\n' +
			'Public Property Set Child(ByVal value As Person)\n' +
			'    Set mChild = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
	});

	it('flags Set used against scalar source-backed members', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-requires-object',
		);
		expectDiagnostic(src, hits, 'set-requires-object', { span: 'Age' });
	});

	it('flags Set assignment between incompatible source-backed object member types', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Class1\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Class1' });
	});

	it('errors on assignment to a read-only class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'readonly-member-assignment',
		);
		expectDiagnostic(src, hits, 'readonly-member-assignment', { span: 'Age' });
	});

	it('errors when a known class receiver uses an unknown member in assignment', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Height = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'Height' });
	});

	it('errors when a known class receiver calls an unknown method', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Delete');
	});

	it('errors when current class Me uses an unknown member', () => {
		const src =
			'Public Sub Save()\n' +
			'    Me.Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Person',
				moduleKind: 'class',
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: src },
				]),
			}),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'Delete' });
	});

	it('accepts known class members and ambiguous project receiver types', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'    p.Unknown = 2\n' +
			'End Sub\n';
		const knownDiagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(knownDiagnostics, 'member-not-found')).toHaveLength(1);
		expect(spanText(src, byCode(knownDiagnostics, 'member-not-found')[0])).toBe(
			'Unknown',
		);
		const ambiguous = projectClassMembers([
			{ moduleName: 'Person', moduleKind: 'class', source: person },
			{ moduleName: 'Other', moduleKind: 'class', source: person.replace(/Age/g, 'Size') },
		]).map((type) => ({ ...type, name: 'Person' }));
		expect(
			byCode(analyzeModule(src, { projectClassMembers: ambiguous }), 'member-not-found'),
		).toHaveLength(0);
	});

	it('does not use source-only document module members to prove absence', () => {
		const workbook =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim wb As ThisWorkbook\n' +
			'    wb.DoesntExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: workbook },
			]),
		});
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('errors when a known standard-module qualifier uses an unknown exported member', () => {
		const caller =
			'Public Sub T()\n' +
			'    Globals.SharedValue = 1\n' +
			'    Globals.MissingValue = 2\n' +
			'    Globals.MissingProcedure\n' +
			'End Sub\n';
		const globals =
			'Public SharedValue As Long\n' +
			'Public Sub Save()\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', moduleKind: 'standard', source: globals },
		], 'Caller');
		expectDiagnostics(caller, diagnostics, 'member-not-found', [
			{ span: 'MissingValue', message: 'Globals.MissingValue' },
			{ span: 'MissingProcedure', message: 'Globals.MissingProcedure' },
		]);
	});

	it('does not expose private standard-module members through qualified member diagnostics', () => {
		const caller =
			'Public Sub T()\n' +
			'    Helpers.Hidden\n' +
			'End Sub\n';
		const helpers =
			'Private Sub Hidden()\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', moduleKind: 'standard', source: helpers },
		], 'Caller');
		expectDiagnostic(caller, diagnostics, 'member-not-found', { span: 'Hidden' });
	});

	it('does not treat unknown external-style qualifiers as project member surfaces', () => {
		const caller =
			'Public Sub T()\n' +
			'    Scripting.Dictionary.CompareMode = TextCompare\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [], 'Caller');
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('errors when ThisWorkbook uses a member absent from source and the exhaustive Workbook host surface', () => {
		const workbook =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.doesnotexist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: workbook },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'doesnotexist' });
	});

	it('does not treat Workbook events as callable object members', () => {
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.AfterSave True\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'AfterSave' });
	});

	it('accepts ThisWorkbook members from source and the exhaustive Workbook host surface', () => {
		const workbook =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.Hello\n' +
			'    ThisWorkbook.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: workbook },
			]),
		});
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('uses the exhaustive Workbook host surface for ActiveWorkbook', () => {
		const src =
			'Public Sub T()\n' +
			'    ActiveWorkbook.doesnotexist\n' +
			'    ActiveWorkbook.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: '' },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'doesnotexist' });
	});

	it('uses the exhaustive Workbook host surface for declared Workbook variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.doesnotexist\n' +
			'    wb.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: '' },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'doesnotexist' });
	});

	it('uses the exhaustive Worksheet host surface for ActiveSheet', () => {
		const src =
			'Public Sub T()\n' +
			'    ActiveSheet.asdf\n' +
			'    ActiveSheet.Range("A1")\n' +
			'    ActiveSheet.Buttons\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'asdf' });
	});

	it('lets local object declarations shadow host globals for member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet As Object\n' +
			'    ActiveSheet.asdf\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('lets untyped local declarations shadow host globals for member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet\n' +
			'    ActiveSheet.asdf\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set refinement for untyped host-global shadows in hard member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet\n' +
			'    Set ActiveSheet = ThisWorkbook\n' +
			'    ActiveSheet.asdf\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('uses local project declarations before host globals for member diagnostics', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet As Person\n' +
			'    ActiveSheet.Range\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);

		expectDiagnostic(src, hits, 'member-not-found', { span: 'Range' });
	});

	it('uses the exhaustive Worksheet host surface through workbook worksheet chains', () => {
		const src =
			'Public Sub T()\n' +
			'    Workbooks(1).Worksheets(1).asdf\n' +
			'    Workbooks(1).Worksheets(1).Range("A1")\n' +
			'    Workbooks(1).Worksheets(1).Buttons\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'asdf' });
	});

	it('uses the exhaustive Range host surface for declared, global, and chained receivers', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim rng As Range\n' +
			'    ActiveCell.NoSuchMember\n' +
			'    ActiveCell.Value2\n' +
			'    rng.DoesNotExist\n' +
			'    rng.ClearContents\n' +
			'    Workbooks(1).Worksheets(1).Range("A1").MissingRangeMember\n' +
			'    Workbooks(1).Worksheets(1).Range("A1").Offset(1, 0).Value = 1\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'member-not-found', [
			{ span: 'NoSuchMember', message: 'Excel.Range.' },
			{ span: 'DoesNotExist', message: 'Excel.Range.' },
			{ span: 'MissingRangeMember', message: 'Excel.Range.' },
		]);
	});

	it('uses the exhaustive Application host surface only after generated promotion', () => {
		const src =
			'Public Sub T()\n' +
			'    Application.DoesNotExist\n' +
			'    Application.CentimetersToPoints 1\n' +
			'    Application.SheetCalculate\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'member-not-found', [
			{ span: 'DoesNotExist', message: 'Excel.Application.DoesNotExist' },
			{ span: 'SheetCalculate', message: 'Excel.Application.SheetCalculate' },
		]);
	});

	it('uses exhaustive generated collection surfaces, including mixed sheet items once all candidates are promoted', () => {
		const src =
			'Public Sub T()\n' +
			'    Workbooks.MissingCollectionMember\n' +
			'    Workbooks.Open "Book.xlsx"\n' +
			'    Worksheets.MissingCollectionMember\n' +
			'    Worksheets.Add\n' +
			'    Sheets.MissingCollectionMember\n' +
			'    Sheets.Add\n' +
			'    Sheets(1).UnknownSheetOrChartMember\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'member-not-found', [
			{ span: 'MissingCollectionMember', message: 'Excel.Workbooks.MissingCollectionMember' },
			{ span: 'MissingCollectionMember', message: 'Excel.Worksheets.MissingCollectionMember' },
			{ span: 'MissingCollectionMember', message: 'Excel.Sheets.MissingCollectionMember' },
			{ span: 'UnknownSheetOrChartMember' },
		]);
	});

	it('uses promoted table, chart, formatting, name, and window host surfaces', () => {
		const src =
			'Public Sub T(ws As Worksheet, rng As Range)\n' +
			'    ws.ListObjects.MissingListObjectsMember\n' +
			'    ws.ListObjects(1).MissingListObjectMember\n' +
			'    ws.ListObjects(1).ListColumns.MissingListColumnsMember\n' +
			'    ws.ListObjects(1).ListRows.MissingListRowsMember\n' +
			'    ws.ChartObjects.MissingChartObjectsMember\n' +
			'    ws.ChartObjects(1).MissingChartObjectMember\n' +
			'    ws.Shapes.MissingShapesMember\n' +
			'    ws.Shapes(1).MissingShapeMember\n' +
			'    rng.Font.MissingFontMember\n' +
			'    rng.Interior.MissingInteriorMember\n' +
			'    rng.Borders.MissingBordersMember\n' +
			'    rng.Borders(1).MissingBorderMember\n' +
			'    rng.FormatConditions.MissingFormatConditionsMember\n' +
			'    rng.FormatConditions.Add(xlCellValue, xlEqual, 1).MissingFormatConditionMember\n' +
			'    rng.Validation.MissingValidationMember\n' +
			'    rng.Hyperlinks.MissingHyperlinksMember\n' +
			'    rng.Hyperlinks(1).MissingHyperlinkMember\n' +
			'    rng.Areas.MissingAreasMember\n' +
			'    rng.Style.MissingStyleMember\n' +
			'    ws.PageSetup.MissingPageSetupMember\n' +
			'    ThisWorkbook.Names.MissingNamesMember\n' +
			'    ThisWorkbook.Names(1).MissingNameMember\n' +
			'    Application.Windows.MissingWindowsMember\n' +
			'    Application.Windows(1).MissingWindowMember\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'MissingListObjectsMember',
			'MissingListObjectMember',
			'MissingListColumnsMember',
			'MissingListRowsMember',
			'MissingChartObjectsMember',
			'MissingChartObjectMember',
			'MissingShapesMember',
			'MissingShapeMember',
			'MissingFontMember',
			'MissingInteriorMember',
			'MissingBordersMember',
			'MissingBorderMember',
			'MissingFormatConditionsMember',
			'MissingFormatConditionMember',
			'MissingValidationMember',
			'MissingHyperlinksMember',
			'MissingHyperlinkMember',
			'MissingAreasMember',
			'MissingStyleMember',
			'MissingPageSetupMember',
			'MissingNamesMember',
			'MissingNameMember',
			'MissingWindowsMember',
			'MissingWindowMember',
		]);
	});

	it('keeps oracle-accepted WorksheetFunction and Pivot unknown members out of hard diagnostics', () => {
		const src =
			'Public Sub T(ws As Worksheet)\n' +
			'    Application.WorksheetFunction.MissingWorksheetFunctionMember\n' +
			'    Application.WorksheetFunction.Sum 1, 2\n' +
			'    ws.PivotTables.MissingPivotTablesMember\n' +
			'    ws.PivotTables(1).MissingPivotTableMember\n' +
			'    ws.PivotTables(1).PivotFields.MissingPivotFieldsMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").MissingPivotFieldMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").PivotItems.MissingPivotItemsMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").PivotItems(1).MissingPivotItemMember\n' +
			'    ws.PivotTables(1).PivotCache.MissingPivotCacheMember\n' +
			'    ws.PivotTables(1).PivotFilters.MissingPivotFiltersMember\n' +
			'    ws.PivotTables(1).PivotFilters(1).MissingPivotFilterMember\n' +
			'    ws.PivotTables(1).CalculatedFields.MissingCalculatedFieldsMember\n' +
			'    ws.PivotTables(1).CalculatedFields.Add("Calc", "=1").MissingCalculatedPivotFieldMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").CalculatedItems.MissingCalculatedItemsMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").CalculatedItems.Add("Q1", "=1").MissingCalculatedPivotItemMember\n' +
			'    ws.PivotTables(1).CubeFields.MissingCubeFieldsMember\n' +
			'    ws.PivotTables(1).CubeFields(1).MissingCubeFieldMember\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(0);
	});

	it('keeps newly promoted metadata-only host families out of hard diagnostics', () => {
		const src =
			'Public Sub T(ws As Worksheet, ch As Chart)\n' +
			'    ws.QueryTables(1).MissingQueryTableMember\n' +
			'    ws.QueryTables(1).Parameters(1).MissingParameterMember\n' +
			'    ThisWorkbook.Connections(1).MissingWorkbookConnectionMember\n' +
			'    ThisWorkbook.Connections(1).OLEDBConnection.MissingOleDbConnectionMember\n' +
			'    ThisWorkbook.Connections(1).ModelTables(1).MissingModelTableMember\n' +
			'    ch.SeriesCollection(1).MissingSeriesMember\n' +
			'    ch.SeriesCollection(1).Points(1).MissingPointMember\n' +
			'    ch.SeriesCollection(1).Trendlines(1).MissingTrendlineMember\n' +
			'    ch.Axes(1).MissingAxisMember\n' +
			'    ch.Axes(1).AxisTitle.MissingAxisTitleMember\n' +
			'    ws.Shapes.Range.MissingShapeRangeMember\n' +
			'    ws.Shapes.Range.GroupItems.MissingGroupShapesMember\n' +
			'    ws.Comments(1).MissingCommentMember\n' +
			'    ws.CommentsThreaded(1).MissingThreadedCommentMember\n' +
			'    ws.AutoFilter.Filters(1).MissingFilterMember\n' +
			'    ws.Sort.SortFields.Add(ws.Range("A1")).MissingSortFieldMember\n' +
			'    ws.OLEObjects(1).MissingOleObjectMember\n' +
			'    ws.Buttons(1).MissingButtonMember\n' +
			'    ws.CheckBoxes(1).MissingCheckBoxMember\n' +
			'    ws.DropDowns(1).MissingDropDownMember\n' +
			'    ws.OptionButtons(1).MissingOptionButtonMember\n' +
			'    ThisWorkbook.SlicerCaches.MissingSlicerCachesMember\n' +
			'    ThisWorkbook.SlicerCaches(1).MissingSlicerCacheMember\n' +
			'    ThisWorkbook.SlicerCaches(1).Slicers.MissingSlicersMember\n' +
			'    ThisWorkbook.SlicerCaches(1).SlicerItems(1).MissingSlicerItemMember\n' +
			'    ThisWorkbook.SlicerCaches(1).TimelineState.MissingTimelineStateMember\n' +
			'    ws.Shapes.Item(1).Fill.MissingFillFormatMember\n' +
			'    ws.Shapes.Item(1).Line.MissingLineFormatMember\n' +
			'    ws.Shapes.Item(1).TextFrame.MissingTextFrameMember\n' +
			'    ws.Shapes.Item(1).TextFrame2.MissingTextFrame2Member\n' +
			'    ws.Shapes.Item(1).PictureFormat.MissingPictureFormatMember\n' +
			'    ws.Shapes.Item(1).Nodes.MissingShapeNodesMember\n' +
			'    ch.ChartTitle.MissingChartTitleMember\n' +
			'    ch.Legend.LegendEntries.MissingLegendEntriesMember\n' +
			'    ch.ChartGroups.MissingChartGroupsMember\n' +
			'    ch.DataTable.MissingDataTableMember\n' +
			'    ch.Walls.MissingWallsMember\n' +
			'    ws.Range("A1").FormatConditions.AddDatabar.MissingDatabarMember\n' +
			'    ws.Range("A1").FormatConditions.AddColorScale(3).ColorScaleCriteria.MissingColorScaleCriteriaMember\n' +
			'    ws.Range("A1").DisplayFormat.MissingDisplayFormatMember\n' +
			'    ws.Drawings.MissingDrawingsMember\n' +
			'    ws.Pictures.MissingPicturesMember\n' +
			'    ws.Lines.MissingLinesMember\n' +
			'    ws.Range("A1").SparklineGroups.MissingSparklineGroupsMember\n' +
			'    ThisWorkbook.XmlMaps.MissingXmlMapsMember\n' +
			'    ThisWorkbook.XmlMaps(1).Schemas.MissingXmlSchemasMember\n' +
			'    ThisWorkbook.PublishObjects.MissingPublishObjectsMember\n' +
			'    ThisWorkbook.WebOptions.MissingWebOptionsMember\n' +
			'    Application.DefaultWebOptions.MissingDefaultWebOptionsMember\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(0);
	});

	it('uses the exhaustive runtime object surface for Err', () => {
		const src =
			'Public Sub T()\n' +
			'    Err.Raise vbObjectError + 1, "M", "boom"\n' +
			'    Err.DoesNotExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'DoesNotExist' });
	});

	it('uses the exhaustive runtime object surface for Debug', () => {
		const src =
			'Public Sub T()\n' +
			'    Debug.Print "value", 1\n' +
			'    Debug.Assert True\n' +
			'    Debug.DoesNotExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-count')).toHaveLength(0);
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'DoesNotExist' });
	});

	it('uses the current workbook Me host surface for ThisWorkbook modules', () => {
		const src =
			'Public Sub T()\n' +
			'    Me.asdf\n' +
			'    Me.AcceptAllChanges\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'asdf' });
	});

	it('uses the exhaustive Worksheet host surface for declared Worksheet variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = ActiveSheet\n' +
			'    ws.asdf\n' +
			'    ws.Range("A1")\n' +
			'    ws.Buttons\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'asdf' });
	});

	it('does not prove missing members from late-bound Object or Variant receivers', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    obj.asdf\n' +
			'    flexible.asdf\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set-assignment refinement for hard late-bound member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Set obj = ActiveSheet\n' +
			'    Set flexible = Workbooks(1).Worksheets(1)\n' +
			'    obj.asdf\n' +
			'    flexible.asdf\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set-assignment refinement for hard late-bound member-call diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Set obj = ActiveSheet\n' +
			'    Set flexible = Workbooks(1).Worksheets(1)\n' +
			'    obj.Range()\n' +
			'    flexible.Range()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('uses an exhaustive host object model to prove a missing member', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					exhaustive: true,
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src =
			'Public Sub T()\n' +
			'    Thing.Missing\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src, { hostModel: model }), 'member-not-found', {
			span: 'Missing',
		});
	});

	it('does not use a curated non-exhaustive host object model to prove absence', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src =
			'Public Sub T()\n' +
			'    Thing.Missing\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { hostModel: model }), 'member-not-found'),
		).toHaveLength(0);
	});
});

describe('analyzeModule - missing Function return assignment', () => {
	it('warns when a Function never assigns its return variable', () => {
		const src =
			'Public Function myFunction()\n' +
			'\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expectDiagnostic(src, hits, 'missing-return-assignment', {
			severity: 'warning',
			span: 'myFunction',
		});
	});

	it('does not warn when a typed Function or Property Get falls through', () => {
		const src =
			'Public Function Label() As String\n' +
			'End Function\n' +
			'\n' +
			'Public Property Get Name() As String\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts scalar and object Function return assignments', () => {
		const src =
			'Public Function Label() As String\n' +
			'    Label = "ready"\n' +
			'End Function\n' +
			'\n' +
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Person\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to known ByRef helper parameters', () => {
		const src =
			'Public Function Run(ParamArray params() As Variant)\n' +
			'    Call CopyVariant(Run, RunEx(params))\n' +
			'End Function\n' +
			'\n' +
			'Private Function RunEx(ByVal values As Variant) As Variant\n' +
			'    RunEx = values\n' +
			'End Function\n' +
			'\n' +
			'Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to project-visible ByRef helper parameters', () => {
		const runSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    Call CopyVariant(dest:=Run, value:=1)\n' +
			'End Function\n';
		const helpersSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			runSrc,
			[
				{ moduleName: 'Runner', source: runSrc },
				{ moduleName: 'Helpers', source: helpersSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to module-qualified ByRef helper parameters', () => {
		const runSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    Helpers.CopyVariant Run, 1\n' +
			'End Function\n';
		const helpersSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			runSrc,
			[
				{ moduleName: 'Runner', source: runSrc },
				{ moduleName: 'Helpers', source: helpersSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(0);
	});

	it('does not count ByVal or ambiguous project helper arguments as return assignments', () => {
		const byValRunSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    CopyVariant Run, 1\n' +
			'End Function\n';
		const byValHelperSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByVal dest As Variant, ByVal value As Variant)\n' +
			'End Sub\n';
		const ambiguousRunSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    CopyVariant Run, 1\n' +
			'End Function\n';
		const byRefHelperSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			byValRunSrc,
			[
				{ moduleName: 'Runner', source: byValRunSrc },
				{ moduleName: 'Helpers', source: byValHelperSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(1);
		expect(byCode(analyzeProjectModule(
			ambiguousRunSrc,
			[
				{ moduleName: 'Runner', source: ambiguousRunSrc },
				{ moduleName: 'HelpersA', source: byRefHelperSrc },
				{ moduleName: 'HelpersB', source: byRefHelperSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(1);
	});

	it('accepts return assignments in the active default VBA7 branch', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			'    HandleValue = 7\n' +
			'#Else\n' +
			'    HandleValue = 6\n' +
			'#End If\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts active VBA7 return assignments with blank lines before #End If', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			' HandleValue = 0\n' +
			'#Else\n' +
			'HandleValue = 1\n' +
			'\n' +
			'\n' +
			'#End If\n' +
			'    \n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('does not count return assignments from inactive default VBA7 branches', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			'    Debug.Print "active branch"\n' +
			'#Else\n' +
			'    HandleValue = 6\n' +
			'#End If\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('HandleValue');
	});

	it('does not warn on conditionally split VBA7 Function headers with a shared body', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Function HandleValue()\n' +
			'#Else\n' +
			'Public Function HandleValue()\n' +
			'#End If\n' +
			'    HandleValue = 1\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('keeps parsing after conditionally split procedure headers with a shared body', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Function Create(ByVal pointer As LongPtr) As Object\n' +
			'#Else\n' +
			'Public Function Create(ByVal pointer As Long) As Object\n' +
			'#End If\n' +
			'    Set Create = Nothing\n' +
			'End Function\n' +
			'Friend Sub Init(ByVal mode As Long)\n' +
			'    Select Case mode\n' +
			'        Case 1\n' +
			'            With Create(0)\n' +
			'                .Name = "ready"\n' +
			'            End With\n' +
			'    End Select\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('does not flag unknown-constant alternate procedure headers as declarations in a procedure', () => {
		const src =
			'#If FULL_INTELLISENSE Then\n' +
			'Public Function AsAcc() As stdAcc\n' +
			'#Else\n' +
			'Public Function AsAcc() As Object\n' +
			'#End If\n' +
			'    Set AsAcc = Nothing\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('ignores exported member Attribute lines while parsing procedure bodies', () => {
		const src =
			'Attribute VB_Name = "Module1"\n' +
			'Friend Sub Init(ByVal mode As Long)\n' +
			'Attribute Init.VB_Description = "Initialises this object."\n' +
			'    Select Case mode\n' +
			'        Case 1\n' +
			'            With Nothing\n' +
			'                .Name = "ready"\n' +
			'            End With\n' +
			'    End Select\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('checks Property Get procedures and ignores Subs', () => {
		const src =
			'Public Property Get Name()\n' +
			'End Property\n' +
			'\n' +
			'Public Sub Refresh()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Name');
	});
});

describe('analyzeModule - string arithmetic coercion', () => {
	it('errors on a nonnumeric string literal inside a numeric assignment expression', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As Integer\n' +
			'    shouldErrorTest1 = 1 + "string"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'string-arithmetic-coercion', {
			severity: 'error',
			span: '"string"',
		});
	});

	it('does not warn on numeric strings or unknown arithmetic operands', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim n As Integer\n' +
			'    n = 1 + "2"\n' +
			'    n = 1 + UnknownValue\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(0);
	});

	it('accepts plus between string literals assigned to a String variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As String\n' +
			'    shouldErrorTest1 = "string" + "string"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - division by zero', () => {
	it('errors on literal zero divisors for division, integer division, and Mod', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'    a = 1 \\ 0\n' +
			'    a = 1 Mod 0\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'division-by-zero', [
			{ span: '0' },
			{ span: '0' },
			{ span: '0' },
		]);
	});

	it('detects parenthesized and signed zero divisors inside nested expressions', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Variant\n' +
			'    a = IIf(True, 1, 1 / (-0))\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('-0');
	});

	it('errors on non-decimal zero literal divisors', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Long\n' +
			'    a = 1 / &H0\n' +
			'    a = 1 \\ &O0\n' +
			'    a = 1 Mod &H1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['&H0', '&O0']);
	});

	it('errors on hex and octal Const zero divisors', () => {
		const src =
			'Private Const ModuleHexZero As Long = &H0\n' +
			'Public Sub T()\n' +
			'    Const LocalOctalZero As Long = &O0\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ModuleHexZero\n' +
			'    a = 1 \\ LocalOctalZero\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleHexZero', 'LocalOctalZero']);
	});

	it('folds hex Const expressions for zero divisor checks', () => {
		const src =
			'Public Sub T()\n' +
			'    Const HexZero As Long = &H1 - &H1\n' +
			'    Const HexOne As Long = &H1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / HexZero\n' +
			'    a = 1 Mod HexOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('HexZero');
	});

	it('errors on zero-valued Enum member divisors', () => {
		const src =
			'Private Enum Denominator\n' +
			'    ExplicitZero = 0\n' +
			'    ImplicitOne\n' +
			'    ExplicitTwo = 2\n' +
			'End Enum\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ExplicitZero\n' +
			'    a = 1 / ImplicitOne\n' +
			'    a = 1 Mod ExplicitTwo\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ExplicitZero');
	});

	it('folds implicit and expression Enum member values for zero divisor checks', () => {
		const src =
			'Public Enum Denominator\n' +
			'    ImplicitZero\n' +
			'    ImplicitOne\n' +
			'    ExpressionZero = ImplicitOne - 1\n' +
			'End Enum\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 \\ ImplicitZero\n' +
			'    a = 1 / ExpressionZero\n' +
			'    a = 1 Mod ImplicitOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ImplicitZero', 'ExpressionZero']);
	});

	it('errors on zero-valued module and local Const divisors', () => {
		const src =
			'Private Const ModuleZero As Long = 0\n' +
			'Private Const ModuleOne As Long = 0 + 1\n' +
			'Public Sub T()\n' +
			'    Const LocalZero As Long = 1 - 1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ModuleZero\n' +
			'    a = 1 / LocalZero\n' +
			'    a = 1 / ModuleOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleZero', 'LocalZero']);
	});

	it('folds visible cross-module Const and Enum values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 / SharedOne\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'    a = 1 \\ SharedOneDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Const SharedOne As Long = 1\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'    SharedOneDivisor = 1\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual(['SharedZero', 'SharedZeroDivisor']);
	});

	it('folds verified runtime and host constants for zero divisor checks', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / vbFalse\n' +
			'    a = 1 / VBA.vbFalse\n' +
			'    a = 1 / xlAbove\n' +
			'    a = 1 / xlNo\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');

		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['vbFalse', 'VBA.vbFalse', 'xlAbove']);
	});

	it('lets source declarations shadow verified runtime and host constants', () => {
		const src =
			'Private Const vbFalse As Long = 1\n' +
			'Public Sub T()\n' +
			'    Dim xlAbove As Long\n' +
			'    Dim a As Double\n' +
			'    xlAbove = 1\n' +
			'    a = 1 / vbFalse\n' +
			'    a = 1 / xlAbove\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(0);
	});

	it('folds module-qualified cross-module Const and Enum values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedDivisors.SharedZero\n' +
			'    a = 1 / SharedDivisors.SharedOne\n' +
			'    a = 1 \\ SharedDivisors.SharedZeroDivisor\n' +
			'    a = 1 \\ SharedDivisors.SharedOneDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Const SharedOne As Long = 1\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'    SharedOneDivisor = 1\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedDivisors.SharedZero',
			'SharedDivisors.SharedZeroDivisor',
		]);
	});

	it('folds exported cross-module Const values through private same-module helpers for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZeroFromHidden\n' +
			'    a = 1 / SharedOneFromHidden\n' +
			'    a = 1 / SharedDivisors.SharedZeroFromHidden\n' +
			'    a = 1 / SharedDivisors.SharedOneFromHidden\n' +
			'    a = 1 / SharedDivisors.HiddenZero\n' +
			'End Sub\n';
		const shared =
			'Private Const HiddenZero As Long = 0\n' +
			'Private Const HiddenOne As Long = 1\n' +
			'Public Const SharedZeroFromHidden As Long = HiddenZero\n' +
			'Public Const SharedOneFromHidden As Long = HiddenOne\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedZeroFromHidden',
			'SharedDivisors.SharedZeroFromHidden',
		]);
	});

	it('defers hidden and ambiguous cross-module Const values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / HiddenZero\n' +
			'    a = 1 / AmbiguousZero\n' +
			'End Sub\n';
		const hidden =
			'Private Const HiddenZero As Long = 0\n';
		const first =
			'Public Const AmbiguousZero As Long = 0\n';
		const second =
			'Public Const AmbiguousZero As Long = 0\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'HiddenDivisors', source: hidden, moduleKind: 'standard' },
				{ moduleName: 'FirstDivisors', source: first, moduleKind: 'standard' },
				{ moduleName: 'SecondDivisors', source: second, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('lets current-module and local Const values shadow cross-module zero divisors', () => {
		const caller =
			'Private Const SharedZero As Long = 1\n' +
			'Public Sub T()\n' +
			'    Const SharedZeroDivisor As Long = 1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('lets source variables shadow cross-module zero divisors', () => {
		const caller =
			'Private SharedZero As Double\n' +
			'Public Sub T()\n' +
			'    Dim SharedZeroDivisor As Double\n' +
			'    Dim a As Double\n' +
			'    SharedZero = 2\n' +
			'    SharedZeroDivisor = 2\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('folds parenthesized Const expressions used as divisors', () => {
		const src =
			'Private Const Zero As Long = 0\n' +
			'Private Const One As Long = 1\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / (Zero + 0)\n' +
			'    a = 1 / (Zero + One)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Zero + 0');
	});

	it('does not flag nonzero literals, variables, or nonzero parenthesized expressions', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    Dim denominator As Double\n' +
			'    a = 1 / 2\n' +
			'    a = 1 / denominator\n' +
			'    a = 1 / (0 + 1)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(0);
	});

	it('ignores literal zero divisors in inactive conditional-compilation branches', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'#If Enabled Then\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { Enabled: false } },
			}), 'division-by-zero'),
		).toHaveLength(0);
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

describe('analyzeModule - Set assignment validation', () => {
	it('flags Set assignment to a known scalar variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest5 As Integer\n' +
			'    Set shouldErrorTest5 = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-requires-object');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('shouldErrorTest5');
	});

	it('flags Set assignment to a String variable even when assigning a New object expression', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Text As String\n' +
			'    Set text = New Collection\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-requires-object');

		expectDiagnostic(src, hits, 'set-requires-object', { span: 'text' });
	});

	it('flags Set assignment to visible exported scalar globals', () => {
		const caller =
			'Public Sub T()\n' +
			'    Set SharedText = New Collection\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'set-requires-object',
		);

		expectDiagnostic(caller, hits, 'set-requires-object', { span: 'SharedText' });
	});

	it('requires Set for visible exported object globals assigned with plain assignment', () => {
		const caller =
			'Public Sub T()\n' +
			'    SharedObject = New Collection\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedObject As Object\n' },
			], 'Caller'),
			'set-required',
		);

		expectDiagnostic(caller, hits, 'set-required', { span: 'SharedObject' });
	});

	it('uses visible exported scalar globals as bare Set RHS value types', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = SharedText\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'assignment-object-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'assignment-object-type-mismatch', { span: 'SharedText' });
	});

	it('uses module-qualified exported scalar globals as Set RHS value types', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = Globals.SharedText\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'assignment-object-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'assignment-object-type-mismatch', { span: 'SharedText' });
	});

	it('uses verified runtime and host constants as Set RHS value types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = vbFalse\n' +
			'    Set target = VBA.vbFalse\n' +
			'    Set target = xlAbove\n' +
			'    Set target = vbNullString\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-object-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-object-type-mismatch', [
			{ span: 'vbFalse', message: 'vbFalse As VbTriState' },
			{ span: 'vbFalse', message: 'VBA.vbFalse As VbTriState' },
			{ span: 'xlAbove', message: 'xlAbove As Constants' },
			{ span: 'vbNullString', message: 'vbNullString As String' },
		]);
	});

	it('lets source declarations shadow verified external constants for Set RHS value types', () => {
		const src =
			'Public Function vbFalse() As Object\n' +
			'End Function\n' +
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Dim xlAbove As Object\n' +
			'    Set target = vbFalse\n' +
			'    Set target = xlAbove\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('uses host globals as Set RHS object value types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim sheet As Worksheet\n' +
			'    Dim book As Workbook\n' +
			'    Set sheet = Application\n' +
			'    Set book = ActiveCell\n' +
			'    Set sheet = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-object-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-object-type-mismatch', [
			{ span: 'Application', message: 'Application As Excel.Application' },
			{ span: 'ActiveCell', message: 'ActiveCell As Excel.Range' },
		]);
	});

	it('allows compatible and generic host global Set assignments', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim app As Application\n' +
			'    Dim item As Object\n' +
			'    Set app = Application\n' +
			'    Set item = ActiveCell\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('lets source declarations shadow host globals for Set RHS object values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim sheet As Worksheet\n' +
			'    Dim Application As Worksheet\n' +
			'    Dim ActiveCell As Worksheet\n' +
			'    Set sheet = Application\n' +
			'    Set sheet = ActiveCell\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('keeps local shadows and ambiguous exported Set RHS globals quiet', () => {
		const shadowCaller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Dim SharedText\n' +
			'    Set target = SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller'), 'assignment-object-type-mismatch')).toHaveLength(0);

		const ambiguousCaller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller'), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('does not guess Set target types for ambiguous visible exported globals', () => {
		const caller =
			'Public Sub T()\n' +
			'    Set SharedText = New Collection\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'set-requires-object')).toHaveLength(0);
	});

	it('does not flag Set assignment to Object, Variant, or unknown object types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Dim sheet As Worksheet\n' +
			'    Set item = Nothing\n' +
			'    Set flexible = Nothing\n' +
			'    Set sheet = Nothing\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'set-requires-object')).toHaveLength(0);
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

describe('analyzeModule - array ReDim', () => {
	it('flags ReDim of a local fixed-size array', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    ReDim Values(1 To 10) As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-array-redim');

		expectDiagnostic(src, hits, 'fixed-array-redim', { severity: 'error', span: 'Values' });
	});

	it('flags ReDim Preserve and module-level fixed-size arrays', () => {
		const src =
			'Private Values(1 To 3) As Long\n' +
			'Sub T()\n' +
			'    ReDim Preserve Values(1 To 10)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-array-redim');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Values');
	});

	it('flags ReDim of scalar variables', () => {
		const src =
			'Private ModuleValue As Long\n' +
			'Sub T()\n' +
			'    Dim Value As Long\n' +
			'    ReDim Value(1 To 10)\n' +
			'    ReDim Preserve ModuleValue(1 To 10)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-redim');

		expectDiagnostics(src, hits, 'scalar-redim', [
			{ severity: 'error', span: 'Value' },
			{ span: 'ModuleValue' },
		]);
	});

	it('uses visible exported scalar and fixed-array globals for ReDim target shapes', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedScalar(1 To 10)\n' +
			'    ReDim SharedFixed(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedScalar As Long\n' +
					'Public SharedFixed(1 To 3) As Long\n',
			},
		], 'Caller');
		const scalarHits = byCode(diagnostics, 'scalar-redim');
		const fixedHits = byCode(diagnostics, 'fixed-array-redim');

		expect(scalarHits).toHaveLength(1);
		expect(spanText(caller, scalarHits[0])).toBe('SharedScalar');
		expect(fixedHits).toHaveLength(1);
		expect(spanText(caller, fixedHits[0])).toBe('SharedFixed');
	});

	it('does not flag visible exported dynamic arrays as invalid ReDim targets', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedValues(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('accepts ReDim Preserve on Variant and implicit Variant targets', () => {
		const src =
			'Public Function RunEx(ByVal vArr As Variant) As Variant\n' +
			'    ReDim Preserve vArr(0 To 29)\n' +
			'    Dim flexible As Variant\n' +
			'    ReDim flexible(0 To 2)\n' +
			'    Dim implicit\n' +
			'    ReDim implicit(0 To 2)\n' +
			'End Function\n';
		const diagnostics = analyzeModule(src);

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('accepts ReDim on visible exported Variant globals', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedValues(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues As Variant\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('lets local dynamic arrays shadow exported invalid ReDim targets', () => {
		const caller =
			'Sub T()\n' +
			'    Dim SharedScalar() As Long\n' +
			'    Dim SharedFixed() As Long\n' +
			'    ReDim SharedScalar(1 To 10)\n' +
			'    ReDim SharedFixed(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedScalar As Long\n' +
					'Public SharedFixed(1 To 3) As Long\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported ReDim targets quiet', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedValue(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue(1 To 3) As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('accepts dynamic arrays and undeclared ReDim targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim dynamicValues() As Long\n' +
			'    ReDim dynamicValues(1 To 10)\n' +
			'    ReDim implicitValues(1 To 10)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
	});

	it('lets a local dynamic array shadow a module fixed-size array', () => {
		const src =
			'Private Values(1 To 3) As Long\n' +
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
	});

	it('lets a local dynamic array shadow a module scalar for ReDim', () => {
		const src =
			'Private Values As Long\n' +
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
	});

	it('ignores fixed-size arrays in inactive conditional branches', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'fixed-array-redim',
			),
		).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'scalar-redim',
			),
		).toHaveLength(0);
	});

	it('uses the active conditional branch when checking ReDim of scalar variables', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'scalar-redim',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'scalar-redim',
			),
		).toHaveLength(0);
	});

	it('flags ReDim dimensions whose explicit lower bound is greater than the upper bound', () => {
		const src =
			'Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Dim grid() As Long\n' +
			'    ReDim values(10 To 1)\n' +
			'    ReDim Preserve values(-1 To -2)\n' +
			'    ReDim grid(1 To 3, 5 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-impossible-bounds');

		expectDiagnostics(src, hits, 'redim-impossible-bounds', [
			{ severity: 'error', span: '10 To 1', message: ["Run-time error '9'", 'dimension 1'] },
			{ span: '-1 To -2' },
			{ span: '5 To 2', message: 'dimension 2' },
		]);
	});

	it('accepts ReDim bounds that are equal, increasing, unknown, upper-only, or inactive', () => {
		const src =
			'Sub T(ByVal first As Long, ByVal last As Long)\n' +
			'    Dim values() As Long\n' +
			'    ReDim values(1 To 1)\n' +
			'    ReDim values(1 To 10)\n' +
			'    ReDim values(first To last)\n' +
			'    ReDim values(10)\n' +
			'#If Win64 Then\n' +
			'    ReDim values(10 To 1)\n' +
			'#End If\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { Win64: false } },
				}),
				'redim-impossible-bounds',
			),
		).toHaveLength(0);
	});

	it('does not add runtime ReDim bounds diagnostics for known invalid ReDim targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim fixed(1 To 3) As Long\n' +
			'    Dim scalar As Long\n' +
			'    ReDim fixed(10 To 1)\n' +
			'    ReDim scalar(10 To 1)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'redim-impossible-bounds')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(1);
	});

	it('flags ReDim Preserve changing a non-final dimension', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 3, 1 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');

		expectDiagnostic(src, hits, 'redim-preserve-dimension-change', {
			severity: 'error',
			span: 'grid',
		});
	});

	it('flags ReDim Preserve changing the known dimension count', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');

		expectDiagnostic(src, hits, 'redim-preserve-dimension-change');
	});

	it('accepts ReDim Preserve changing only the last dimension', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 2, 1 To 3)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'redim-preserve-dimension-change')).toHaveLength(0);
	});

	it('flags ReDim Preserve changing the lower bound of the final dimension', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 2, 0 To 3)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');

		expectDiagnostic(src, hits, 'redim-preserve-dimension-change');
	});

	it('does not let ReDim shapes learned inside nested blocks leak outward', () => {
		const src =
			'Sub T(ByVal flag As Boolean)\n' +
			'    Dim grid() As Long\n' +
			'    If flag Then\n' +
			'        ReDim grid(1 To 2, 1 To 2)\n' +
			'    End If\n' +
			'    ReDim Preserve grid(1 To 3, 1 To 2)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'redim-preserve-dimension-change')).toHaveLength(0);
	});
});

describe('analyzeModule - unallocated dynamic array access', () => {
	it('flags straight-line local dynamic array indexed before ReDim', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Debug.Print values(0)\n' +
			'    values(1) = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expectDiagnostics(src, hits, 'unallocated-dynamic-array-access', [
			{ severity: 'error', span: 'values' },
			{ span: 'values' },
		]);
	});

	it('accepts allocated and fixed arrays but flags again after Erase', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Dim fixedValues(0 To 1) As Long\n' +
			'    ReDim values(0 To 2)\n' +
			'    Debug.Print values(0)\n' +
			'    Debug.Print fixedValues(0)\n' +
			'    Erase values\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('values');
	});

	it('flags LBound and UBound on unallocated local dynamic arrays', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Debug.Print LBound(values)\n' +
			'    Debug.Print UBound(values, 1)\n' +
			'    Debug.Print VBA.LBound(values)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expectDiagnostics(src, hits, 'unallocated-dynamic-array-access', [
			{ span: 'values', message: 'LBound' },
			{ span: 'values', message: 'UBound' },
			{ span: 'values' },
		]);
	});

	it('accepts bound intrinsics after ReDim and non-intrinsic member calls', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    ReDim values(0 To 2)\n' +
			'    Debug.Print LBound(values)\n' +
			'    Debug.Print Helpers.LBound(values)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('keeps module-level arrays parameters Static locals and helper-touched arrays quiet', () => {
		const src =
			'Private moduleValues() As Long\n' +
			'Private Sub Fill(ByRef target() As Long)\n' +
			'End Sub\n' +
			'Public Sub T(ByRef paramValues() As Long)\n' +
			'    Static cached() As Long\n' +
			'    Dim helperValues() As Long\n' +
			'    Fill helperValues\n' +
			'    Debug.Print moduleValues(0)\n' +
			'    Debug.Print paramValues(0)\n' +
			'    Debug.Print cached(0)\n' +
			'    Debug.Print helperValues(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('stays quiet after nested block allocation makes straight-line state unknown', () => {
		const src =
			'Public Sub T(ByVal ready As Boolean)\n' +
			'    Dim values() As Long\n' +
			'    If ready Then\n' +
			'        ReDim values(0 To 1)\n' +
			'    End If\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('ignores inactive conditional-compilation indexed access', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'#If Enabled Then\n' +
			'    Debug.Print values(0)\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});
});

describe('analyzeModule - array Erase', () => {
	it('flags an Erase arithmetic expression', () => {
		const src =
			'Sub T()\n' +
			'    Erase 1 + 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-erase-target');

		expectDiagnostic(src, hits, 'invalid-erase-target', { severity: 'error', span: '1 + 2' });
	});

	it('flags expression targets in comma-separated Erase lists', () => {
		const src =
			'Sub T()\n' +
			'    Erase Values, other + 1, "bad"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-erase-target');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['other + 1', '"bad"']);
	});

	it('accepts variable-like Erase targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    Erase Values\n' +
			'    Erase Values, OtherValues\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-erase-target')).toHaveLength(0);
	});

	it('flags Erase of declared non-Variant scalar variables', () => {
		const src =
			'Private ModuleValue As Long\n' +
			'Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim Value As Long\n' +
			'    Erase obj, Value, ModuleValue\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'erase-requires-array');

		expectDiagnostics(src, hits, 'erase-requires-array', [
			{ severity: 'error', span: 'obj' },
			{ span: 'Value' },
			{ span: 'ModuleValue' },
		]);
	});

	it('uses visible exported scalar globals for Erase target shapes', () => {
		const caller =
			'Sub T()\n' +
			'    Erase SharedValue, SharedObject\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValue As Long\n' +
					'Public SharedObject As Object\n',
			},
		], 'Caller');
		const hits = byCode(diagnostics, 'erase-requires-array');

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual(['SharedValue', 'SharedObject']);
	});

	it('accepts Erase of visible exported arrays and Variant globals', () => {
		const caller =
			'Sub T()\n' +
			'    Erase SharedValues, SharedFlexible\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValues() As Long\n' +
					'Public SharedFlexible As Variant\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'erase-requires-array')).toHaveLength(0);
	});

	it('lets local array and Variant declarations shadow exported Erase scalars', () => {
		const caller =
			'Sub T()\n' +
			'    Dim SharedValue() As Long\n' +
			'    Dim SharedObject As Variant\n' +
			'    Erase SharedValue, SharedObject\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValue As Long\n' +
					'Public SharedObject As Object\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'erase-requires-array')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported Erase targets quiet', () => {
		const caller =
			'Sub T()\n' +
			'    Erase SharedValue\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue As Object\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'erase-requires-array')).toHaveLength(0);
	});

	it('accepts Erase of arrays, Variants, unresolved names, and non-simple targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    Dim FixedValues(1 To 3) As Long\n' +
			'    Dim Flexible As Variant\n' +
			'    Dim ImplicitVariant\n' +
			'    Erase Values, FixedValues, Flexible, ImplicitVariant, UnknownValues\n' +
			'    Erase Settings.Values, Values(1)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'erase-requires-array')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking Erase scalar targets', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    Erase Values\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'erase-requires-array',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'erase-requires-array',
			),
		).toHaveLength(0);
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

describe('analyzeModule - object module public declaration restrictions', () => {
	it('flags public declarations that cannot be object-module members', () => {
		const src =
			'Public Const MaxRows As Long = 1000\n' +
			'Public Names() As String\n' +
			'Public FixedName As String * 20\n' +
			'Public Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
			'object-module-public-member',
		);
		expectDiagnostics(src, hits, 'object-module-public-member', [
			{ severity: 'error', span: 'MaxRows' },
			{ span: 'Names' },
			{ span: 'FixedName' },
			{ span: 'Customer' },
			{ span: 'Sleep' },
		]);
	});

	it('does not apply object-module public-member restrictions in standard modules', () => {
		const src =
			'Public Const MaxRows As Long = 1000\n' +
			'Public Names() As String\n' +
			'Public FixedName As String * 20\n' +
			'Public Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		expect(byCode(analyzeModule(src), 'object-module-public-member')).toHaveLength(0);
	});

	it('does not flag private object-module declarations in this public-member rule', () => {
		const src =
			'Private Const MaxRows As Long = 1000\n' +
			'Private Names() As String\n' +
			'Private FixedName As String * 20\n' +
			'Private Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
				'object-module-public-member',
			),
		).toHaveLength(0);
	});

	it('accepts private fixed-length strings across object module kinds', () => {
		const src = 'Private FixedName As String * 20\n';
		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			const diagnostics = analyzeModule(src, { moduleName: 'ObjectModule', moduleKind });
			expect(byCode(diagnostics, 'object-module-public-member')).toHaveLength(0);
			expect(byCode(diagnostics, 'fixed-length-string-size')).toHaveLength(0);
		}
	});
});

describe('analyzeModule - Event declaration module-kind restrictions', () => {
	it('flags Event declarations in standard modules', () => {
		const src =
			'Public Event BeforeAdd(ByRef arr As Variant, ByRef cancel As Boolean)\n' +
			'Private Event AfterAdd(ByRef arr As Variant)\n';
		expectDiagnostics(src, analyzeModule(src), 'event-declaration-module-kind', [
			{ severity: 'error', span: 'BeforeAdd' },
			{ span: 'AfterAdd' },
		]);
	});

	it('ignores inactive standard-module Event declarations', () => {
		const src =
			'#If VBA7 Then\n' +
			'#Else\n' +
			'Public Event LegacyOnly()\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'event-declaration-module-kind')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'event-declaration-module-kind',
			),
		).toHaveLength(1);
	});
});

describe('analyzeModule - WithEvents declaration restrictions', () => {
	it('accepts module-level WithEvents declarations in object modules', () => {
		const src = 'Private WithEvents App As Application\n';
		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			expect(
				byCode(analyzeModule(src, { moduleName: 'EventSource', moduleKind }), 'withevents-declaration'),
			).toHaveLength(0);
		}
	});

	it('flags module-level WithEvents declarations in standard modules', () => {
		const src = 'Private WithEvents App As Application\n';
		expectDiagnostic(src, analyzeModule(src), 'withevents-declaration', { span: 'App' });
	});

	it('flags local WithEvents declarations inside procedures', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim WithEvents App As Application\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'EventSource', moduleKind: 'class' }),
			'withevents-declaration',
		);
		expectDiagnostic(src, hits, 'withevents-declaration', { span: 'App' });
	});

	it('flags WithEvents arrays and As New declarations', () => {
		const src =
			'Private WithEvents App As New Application\n' +
			'Private WithEvents Apps(1 To 2) As Application\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'EventSource', moduleKind: 'class' }),
			'withevents-declaration',
		);
		expectDiagnostics(src, hits, 'withevents-declaration', [
			{ span: 'App', message: 'As New' },
			{ span: 'Apps', message: 'array' },
		]);
	});

	it('ignores inactive WithEvents declarations', () => {
		const src =
			'#If VBA7 Then\n' +
			'#Else\n' +
			'Private WithEvents LegacyApp As Application\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'withevents-declaration')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'withevents-declaration',
			),
		).toHaveLength(1);
	});
});

describe('analyzeModule - Friend declaration restrictions', () => {
	it('accepts Friend procedures in object modules', () => {
		const src = 'Friend Sub InternalOnly()\nEnd Sub\n';
		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			expect(
				byCode(
					analyzeModule(src, { moduleName: 'EventSource', moduleKind }),
					'friend-declaration',
				),
			).toHaveLength(0);
		}
	});

	it('flags Friend procedures in standard modules', () => {
		const src = 'Friend Sub InternalOnly()\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'friend-declaration', {
			severity: 'error',
			span: 'Friend',
		});
	});

	it('flags Friend variable declarations even in object modules', () => {
		const src = 'Friend mValue As Long\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'EventSource', moduleKind: 'class' }),
			'friend-declaration',
		);
		expectDiagnostic(src, hits, 'friend-declaration', { span: 'Friend' });
	});

	it('ignores inactive Friend declarations', () => {
		const src =
			'#If VBA7 Then\n' +
			'#Else\n' +
			'Friend Sub LegacyOnly()\n' +
			'End Sub\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'friend-declaration')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'friend-declaration',
			),
		).toHaveLength(1);
	});
});

describe('analyzeModule - Implements statement placement', () => {
	it('accepts module-level Implements statements in object-module declaration sections', () => {
		const src = 'Option Explicit\nImplements Person\nPrivate Value As Long\n';
		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			expect(
				byCode(
					analyzeModule(src, { moduleName: 'EventSource', moduleKind }),
					'implements-statement-placement',
				),
			).toHaveLength(0);
		}
	});

	it('flags module-level Implements statements in standard modules', () => {
		const src = 'Option Explicit\nImplements Person\n';
		expectDiagnostic(src, analyzeModule(src), 'implements-statement-placement', {
			severity: 'error',
			span: 'Person',
		});
	});

	it('flags module-level Implements statements after procedures in object modules', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub Demo()\n' +
			'End Sub\n' +
			'Implements Person\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'EventSource', moduleKind: 'class' }),
			'implements-statement-placement',
		);
		expectDiagnostic(src, hits, 'implements-statement-placement', { span: 'Person' });
	});

	it('flags Implements statements inside procedure bodies', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub Demo()\n' +
			'    Implements Person\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			moduleName: 'EventSource',
			moduleKind: 'class',
			knownIdentifiers: new Set<string>(),
		});
		expectDiagnostic(src, diagnostics, 'implements-statement-placement', { span: 'Person' });
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('ignores inactive Implements statements', () => {
		const src =
			'#If VBA7 Then\n' +
			'#Else\n' +
			'Implements LegacyInterface\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'implements-statement-placement')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'implements-statement-placement',
			),
		).toHaveLength(1);
	});

	it('reports the full qualified interface name after line labels', () => {
		const src = '10 Implements Excel.Worksheet\n';
		expectDiagnostic(src, analyzeModule(src), 'implements-statement-placement', {
			span: 'Excel.Worksheet',
		});
	});
});

describe('analyzeModule - RaiseEvent target binding', () => {
	it('accepts RaiseEvent statements that target active same-module Event declarations', () => {
		const src =
			'Option Explicit\n' +
			'Public Event Changed(ByVal value As Long)\n' +
			'Private Event Hidden()\n' +
			'Public Sub Touch()\n' +
			'    RaiseEvent Changed(1)\n' +
			'    RaiseEvent Hidden\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'EventSource', moduleKind: 'class' }),
				'raiseevent-undeclared-event',
			),
		).toHaveLength(0);
	});

	it('flags RaiseEvent statements whose event is not declared in the same module', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub Touch()\n' +
			'    RaiseEvent Changed()\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			moduleName: 'EventSource',
			moduleKind: 'class',
			knownIdentifiers: new Set<string>(),
		});
		expectDiagnostic(src, diagnostics, 'raiseevent-undeclared-event', {
			severity: 'error',
			span: 'Changed',
		});
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('uses conditional-compilation activity for Event declarations and RaiseEvent statements', () => {
		const src =
			'#If VBA7 Then\n' +
			'#Else\n' +
			'Public Event LegacyOnly()\n' +
			'#End If\n' +
			'Public Sub Touch()\n' +
			'    RaiseEvent LegacyOnly\n' +
			'    #If VBA7 Then\n' +
			'    #Else\n' +
			'    RaiseEvent MissingWhenVba7\n' +
			'    #End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'raiseevent-undeclared-event')).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'raiseevent-undeclared-event',
			),
		).toHaveLength(1);
	});

	it('reports event names after line labels and leaves partial RaiseEvent quiet', () => {
		const src =
			'Public Sub Touch()\n' +
			'10 RaiseEvent Missing\n' +
			'20 RaiseEvent\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'EventSource', moduleKind: 'class' }),
			'raiseevent-undeclared-event',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Missing');
	});
});

describe('analyzeModule - conditional Declare platform rules', () => {
	it('requires PtrSafe only when the supplied compiler constants prove Win64', () => {
		const src = 'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';

		const hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { Win64: true } },
			}),
			'declare-missing-ptrsafe',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Sleep');

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { Win64: false } },
				}),
				'declare-missing-ptrsafe',
			),
		).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'declare-missing-ptrsafe')).toHaveLength(0);
	});

	it('does not flag inactive legacy Declare branches under Win64', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'#Else\n' +
			'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
			'#End If\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true, Win64: true } },
				}),
				'declare-missing-ptrsafe',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - event handler module scope guidance', () => {
	it('guides when a workbook handler is declared in a standard module', () => {
		const src = 'Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Module1', moduleKind: 'standard' }),
			'event-handler-module-scope',
		);
		expectDiagnostic(src, hits, 'event-handler-module-scope', {
			severity: 'information',
			span: 'Workbook_Open',
		});
	});

	it('does not guide for workbook handlers in ThisWorkbook', () => {
		const src = 'Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
				'event-handler-module-scope',
			),
		).toHaveLength(0);
	});

	it('does not guide for worksheet handlers in worksheet document modules', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'Sheet1', moduleKind: 'document' }),
				'event-handler-module-scope',
			),
		).toHaveLength(0);
	});

	it('guides when a worksheet handler is declared in ThisWorkbook', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
			'event-handler-module-scope',
		);
		expectDiagnostic(src, hits, 'event-handler-module-scope', { span: 'Worksheet_Change' });
	});

	it('uses proven chart document subtype before giving guidance', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Worksheet_Calculate()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'RevenueChart',
				moduleKind: 'document',
				documentType: 'chart',
			}),
			'event-handler-module-scope',
		);
		expectDiagnostic(src, hits, 'event-handler-module-scope', { span: 'Worksheet_Calculate' });
	});

	it('does not guide for chart handlers in chart document modules', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Chart_Calculate()\nEnd Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					moduleName: 'RevenueChart',
					moduleKind: 'document',
					documentType: 'chart',
				}),
				'event-handler-module-scope',
			),
		).toHaveLength(0);
	});

	it('guides when a chart handler is declared in a worksheet document module', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Chart_Calculate()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Sheet1',
				moduleKind: 'document',
				documentType: 'worksheet',
			}),
			'event-handler-module-scope',
		);
		expectDiagnostic(src, hits, 'event-handler-module-scope', { span: 'Chart_Calculate' });
	});
});

describe('analyzeModule - Call requires parentheses', () => {
	it('flags an unparenthesised Call argument list', () => {
		const src = 'Sub T()\n    Call MsgBox "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"hello"');
		expect(hits[0].severity).toBe('error');
	});

	it('flags an unparenthesised member Call', () => {
		const src = 'Sub T()\n    Call obj.Method 1, 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(1);
	});

	it('flags a standalone zero-argument member call that uses empty parentheses without Call', () => {
		const src = 'Sub T()\n    ThisWorkbook.CanCheckIn()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('CanCheckIn()');
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a standalone zero-argument host method call with empty parentheses before arity checks', () => {
		const src = 'Sub T()\n    Application.Calculate()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a standalone zero-argument class method call with empty parentheses', () => {
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    p.Save()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save()');
	});

	it('flags a same-module zero-argument Function call statement with empty parentheses', () => {
		const src =
			'Sub mySub()\n' +
			'    myFunction()\n' +
			'End Sub\n' +
			'\n' +
			'Function myFunction() As String\n' +
			'    myFunction = "hello world!"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('myFunction()');
	});

	it('leaves required-argument procedure empty calls to arity diagnostics', () => {
		const src =
			'Sub mySub()\n' +
			'    Greet()\n' +
			'End Sub\n' +
			'\n' +
			'Sub Greet(ByVal message As String)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Greet');
	});

	it('flags a standalone zero-argument runtime call with empty parentheses', () => {
		const src = 'Sub T()\n    DoEvents()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expectDiagnostic(src, hits, 'call-statement-forbids-parens', { span: 'DoEvents()' });
		expect(hits[0].message).not.toContain('prefixed with Call');
	});

	it('flags DoEvents as an invalid explicit Call target', () => {
		const src =
			'Sub T()\n' +
			'    Call DoEvents\n' +
			'    Call DoEvents()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-explicit-call-target');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DoEvents', 'DoEvents']);
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not apply runtime-only call syntax rules to local-shadowed runtime names', () => {
		const src =
			'Sub T()\n' +
			'    Dim Erl\n' +
			'    Erl()\n' +
			'    Call Erl()\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(diagnostics, 'invalid-explicit-call-target')).toHaveLength(0);
	});

	it('flags an unqualified same-class method call statement with empty parentheses', () => {
		const src =
			'Public Sub SaveAll()\n' +
			'    Save()\n' +
			'End Sub\n' +
			'\n' +
			'Public Sub Save()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
			'call-statement-forbids-parens',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save()');
	});

	it('flags an exported cross-module zero-argument Function call statement with empty parentheses', () => {
		const src =
			'Sub mySub()\n' +
			'    myFunction()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(src, [
				{ moduleName: 'Helpers', source: 'Public Function myFunction() As String\nEnd Function\n' },
			], 'Caller'),
			'call-statement-forbids-parens',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('myFunction()');
	});

	it('validates non-empty standalone member call parentheses with known signatures', () => {
		const src = 'Sub T()\n    Application.Calculate(1)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(1);
	});

	it('accepts Call with parentheses', () => {
		const src = 'Sub T()\n    Call MsgBox("hello")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parameterless Call', () => {
		const src = 'Sub T()\n    Call DoWork\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts a Call to a parenthesised member chain', () => {
		const src = 'Sub T()\n    Call obj.Method(1, 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts parenthesized member calls with Call or in expression context', () => {
		const src =
			'Sub T()\n' +
			'    Dim ok As Boolean\n' +
			'    Call ThisWorkbook.CanCheckIn()\n' +
			'    ok = ThisWorkbook.CanCheckIn()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('accepts known Function calls with Call or in expression context', () => {
		const src =
			'Sub mySub()\n' +
			'    Dim value As String\n' +
			'    Call myFunction()\n' +
			'    value = myFunction()\n' +
			'End Sub\n' +
			'\n' +
			'Function myFunction() As String\n' +
			'    myFunction = "hello world!"\n' +
			'End Function\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('accepts zero-argument runtime calls as bare statements or in expression context', () => {
		const src =
			'Sub T()\n' +
			'    Dim value As Integer\n' +
			'    DoEvents\n' +
			'    value = DoEvents()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'invalid-explicit-call-target')).toHaveLength(0);
	});

	it('does not flag unknown bare empty-parentheses statements as project procedure calls', () => {
		const src = 'Sub T()\n    MaybeExternal()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not flag a parenless non-Call statement', () => {
		const src = 'Sub T()\n    MsgBox "hello"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - expression call requires parentheses', () => {
	it('flags a same-module Function called with parenless arguments in an assignment', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'expression-call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('InvoiceTotal');
	});

	it('flags a runtime Function called with parenless arguments in an assignment', () => {
		const src = 'Sub T()\n    answer = MsgBox "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'expression-call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MsgBox');
	});

	it('does not require expression-call parentheses for local-shadowed runtime names', () => {
		const src =
			'Sub T()\n' +
			'    Dim Format\n' +
			'    answer = Format 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('flags a unique exported project Function called with parenless arguments in an assignment', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'expression-call-requires-parens',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('InvoiceTotal');
	});

	it('flags a module-qualified project Function called with parenless arguments in an assignment', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = Invoices.InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'expression-call-requires-parens',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('InvoiceTotal');
	});

	it('does not guess when a bare exported project Function name is ambiguous', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const first =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const second =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n';

		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'Invoices', source: first },
			{ moduleName: 'AlternateInvoices', source: second },
		], 'Caller'), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('does not treat exported project Subs as expression Functions', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = PrintTotal 100\n' +
			'End Sub\n';
		const helpers = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';

		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', source: helpers },
		], 'Caller'), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parenthesized Function call in an assignment', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal(100, 0.08)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parameterless Function reference in an expression', () => {
		const src =
			'Public Function CurrentTotal() As Currency\nEnd Function\n' +
			'Public Sub TestTotal()\n' +
			'    total = CurrentTotal + 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('accepts parameterless property reads followed by keyword operators', () => {
		const src =
			'Public Property Get Style() As Long\nEnd Property\n' +
			'Public Sub T()\n' +
			'    Dim resizable As Boolean\n' +
			'    resizable = Style And &H40000\n' +
			'    resizable = Style Or &H40000\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid expression syntax', () => {
	it('flags an impossible operator sequence in a call argument expression', () => {
		const src =
			'Sub T()\n' +
			'    MsgBox myFunctionTest***\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('***');
	});

	it('flags unsupported C-style ternary syntax', () => {
		const src = 'Sub T()\n    value = flag ? 1 : 2\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'invalid-expression-syntax', { span: '?' });
	});

	it('flags a statement that ends with a binary operator', () => {
		const src = 'Sub T()\n    total = subtotal *\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('*');
	});

	it('flags trailing member-access dots on object receivers', () => {
		const src = 'Sub T()\n    ThisWorkbook.\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'invalid-expression-syntax', { span: '.' });
	});

	it('flags a bare leading dot inside With as incomplete final source', () => {
		const src =
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('accepts complete leading-dot member access inside With', () => {
		const src =
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .Range("A1").Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('lets scalar member access own known scalar trailing dots', () => {
		const src = 'Sub T()\n    Dim value As String\n    value.\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(1);
	});

	it('lets scalar member access own visible exported scalar trailing dots', () => {
		const caller = 'Sub T()\n    SharedText.\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'invalid-expression-syntax')).toHaveLength(0);
		const hits = byCode(diagnostics, 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedText.');
	});

	it('uses generic incomplete member syntax for local shadows of exported scalars', () => {
		const caller =
			'Sub T()\n' +
			'    Dim SharedText\n' +
			'    SharedText.\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-member-access')).toHaveLength(0);
		const hits = byCode(diagnostics, 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('.');
	});

	it('treats fixed-length String declarations as scalar String receivers', () => {
		const src = 'Sub T()\n    Dim value As String * 20\n    value.\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
	});

	it('does not flag valid arithmetic or string expressions', () => {
		const src =
			'Sub T()\n' +
			'    total = subtotal * taxRate\n' +
			'    message = prefix & suffix\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('accepts IIf as the supported inline conditional function', () => {
		const src = 'Sub T()\n    value = IIf(flag, 1, 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
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

describe('analyzeModule - Exit statement matches procedure', () => {
	it('flags Exit Function inside a Sub', () => {
		const src = 'Sub T()\n    Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
		expect(hits[0].severity).toBe('error');
	});

	it('flags mismatched Exit statements after numeric line labels', () => {
		const src = 'Sub T()\n10 Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
	});

	it('flags Exit Sub inside a Function', () => {
		const src = 'Function F() As Long\n    Exit Sub\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('flags Exit Sub inside a Property', () => {
		const src = 'Property Get Name() As String\n    Exit Sub\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('accepts a matching Exit Sub', () => {
		const src = 'Sub T()\n    Exit Sub\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('accepts a matching Exit Property', () => {
		const src = 'Property Let Name(v As String)\n    Exit Property\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('ignores Exit For and Exit Do', () => {
		const src =
			'Sub T()\n' +
			'    Do\n        Exit Do\n    Loop\n' +
			'    For i = 1 To 3\n        Exit For\n    Next i\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});
});

describe('analyzeModule - procedure labels', () => {
	it('accepts forward and backward GoTo and GoSub labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'    GoTo done\n' +
			'start:\n' +
			'    GoSub done\n' +
			'    GoTo start\n' +
			'done:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('flags missing GoTo, GoSub, Resume, and On Error labels', () => {
		const src =
			'Sub T()\n' +
			'    GoTo MissingGoTo\n' +
			'    GoSub MissingGoSub\n' +
			'    Resume MissingResume\n' +
			'    On Error GoTo MissingHandler\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'MissingGoTo',
			'MissingGoSub',
			'MissingResume',
			'MissingHandler',
		]);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('keeps labels scoped to their enclosing procedure', () => {
		const src =
			'Sub A()\n' +
			'    GoTo Done\n' +
			'End Sub\n' +
			'Sub B()\n' +
			'Done:\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Done');
	});

	it('flags duplicate named labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'StartHere:\n' +
			'    Debug.Print "first"\n' +
			'StartHere:\n' +
			'    Debug.Print "second"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-label');

		expectDiagnostic(src, hits, 'duplicate-label', { severity: 'error', span: 'StartHere' });
	});

	it('flags duplicate normalized numeric labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'010 Debug.Print "first"\n' +
			'10 Debug.Print "second"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-label');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('10');
	});

	it('allows the same label name in separate procedures', () => {
		const src =
			'Sub A()\n' +
			'Done:\n' +
			'End Sub\n' +
			'Sub B()\n' +
			'Done:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-label')).toHaveLength(0);
	});

	it('accepts explicit On Error forms that do not name a label', () => {
		const src =
			'Sub T()\n' +
			'    On Error Resume Next\n' +
			'    On Error GoTo 0\n' +
			'    On Error GoTo -1\n' +
			'    Resume\n' +
			'    Resume Next\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('accepts On Error disable forms after line numbers and inside single-line If statements', () => {
		const src =
			'Sub T(ByVal flag As Boolean)\n' +
			'10 On Error GoTo 0\n' +
			'20 If flag Then On Error GoTo 0 Else On Error GoTo Handler\n' +
			'30 On Error GoTo -1\n' +
			'Handler:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('validates On n GoTo and On n GoSub label lists', () => {
		const src =
			'Sub T(ByVal n As Long)\n' +
			'    On n GoTo First, MissingJump\n' +
			'    On n GoSub First, MissingSub\n' +
			'First:\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['MissingJump', 'MissingSub']);
	});

	it('validates labels on both sides of a single-line If Else statement', () => {
		const src = 'Sub T(ByVal flag As Boolean)\n    If flag Then GoTo MissingA Else GoTo MissingB\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['MissingA', 'MissingB']);
	});

	it('accepts numeric line labels and references', () => {
		const src =
			'Sub T()\n' +
			'    GoTo 10\n' +
			'10:\n' +
			'    On 1 GoTo 10\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('does not use labels from inactive conditional-compilation branches', () => {
		const src =
			'Sub T()\n' +
			'    GoTo MissingWhenInactive\n' +
			'#If VBA7 Then\n' +
			'MissingWhenInactive:\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: false } },
			}),
			'undefined-label',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MissingWhenInactive');
	});

	it('does not treat inactive duplicate labels as active duplicates', () => {
		const src =
			'Sub T()\n' +
			'StartHere:\n' +
			'#If VBA7 Then\n' +
			'StartHere:\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-label',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - statement context', () => {
	it('flags #ElseIf after #Else in the same conditional-compilation block', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#Else\n' +
			'    Debug.Print 1\n' +
			'#ElseIf True Then\n' +
			'    Debug.Print 2\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { severity: 'error', span: '#ElseIf' });
	});

	it('flags duplicate #Else branches in a conditional-compilation block', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#Else\n' +
			'    Debug.Print 1\n' +
			'#Else\n' +
			'    Debug.Print 2\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { span: '#Else' });
	});

	it('accepts conditional-compilation #ElseIf branches before #Else and in nested blocks', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#ElseIf True Then\n' +
			'    Debug.Print 1\n' +
			'#Else\n' +
			'    #If True Then\n' +
			'        Debug.Print 2\n' +
			'    #ElseIf False Then\n' +
			'        Debug.Print 3\n' +
			'    #End If\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'else-branch-order')).toHaveLength(0);
	});

	it('flags ElseIf after Else in the same If block', () => {
		const src =
			'Sub T()\n' +
			'    If False Then\n' +
			'        Debug.Print 0\n' +
			'    Else\n' +
			'        Debug.Print 1\n' +
			'    ElseIf True Then\n' +
			'        Debug.Print 2\n' +
			'    End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { span: 'ElseIf' });
	});

	it('flags duplicate Else branches in the same If block', () => {
		const src =
			'Sub T()\n' +
			'    If False Then\n' +
			'        Debug.Print 0\n' +
			'    Else\n' +
			'        Debug.Print 1\n' +
			'    Else\n' +
			'        Debug.Print 2\n' +
			'    End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { span: 'Else' });
	});

	it('accepts ElseIf branches before Else and nested If blocks after Else', () => {
		const src =
			'Sub T()\n' +
			'    If False Then\n' +
			'        Debug.Print 0\n' +
			'    ElseIf True Then\n' +
			'        Debug.Print 1\n' +
			'    Else\n' +
			'        If True Then\n' +
			'            Debug.Print 2\n' +
			'        ElseIf False Then\n' +
			'            Debug.Print 3\n' +
			'        End If\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'else-branch-order')).toHaveLength(0);
	});

	it('does not report inactive If branch-order diagnostics', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    If False Then\n' +
			'    Else\n' +
			'    ElseIf True Then\n' +
			'    End If\n' +
			'#End If\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'else-branch-order',
			),
		).toHaveLength(0);
	});

	it('flags an If statement missing Then', () => {
		const src = 'Sub T()\n    If x > 0\n        x = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'if-missing-then');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('If');
	});

	it('accepts multiline and single-line If statements with Then', () => {
		const src =
			'Sub T()\n' +
			'    If x > 0 Then\n        x = 1\n    End If\n' +
			'    If x > 1 Then x = 2\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'if-missing-then')).toHaveLength(0);
	});

	it('flags Case outside Select Case', () => {
		const src = 'Sub T()\n    Case 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'case-outside-select');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Case');
	});

	it('accepts Case inside Select Case', () => {
		const src =
			'Sub T()\n' +
			'    Select Case x\n' +
			'        Case 1, 2\n' +
			'            x = 3\n' +
			'        Case Else\n' +
			'            x = 4\n' +
			'    End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
	});

	it('accepts line-numbered Case statements inside Select Case', () => {
		const src =
			'Sub T()\n' +
			'10 Select Case x\n' +
			'20 Case 1, 2\n' +
			'30 x = 3\n' +
			'40 Case Else\n' +
			'50 x = 4\n' +
			'60 End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
	});

	it('flags leading-dot member access outside With', () => {
		const src = 'Sub T()\n    .Value = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'member-access-outside-with');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('accepts leading-dot member access inside With', () => {
		const src =
			'Sub T()\n' +
			'    With Range("A1")\n' +
			'        .Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
	});

	it('uses the With receiver for unknown source-backed class members', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Delete\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'Delete' });
	});

	it('uses the With receiver for source-backed class member assignment rules', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Age = 2\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'readonly-member-assignment',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Age');
	});

	it('uses the With receiver for source-backed class member call rules', () => {
		const person = 'Public Sub Save(ByVal Count As Long)\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Save "bad"\n' +
			'        .Save()\n' +
			'    End With\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(1);
		expect(byCode(diagnostics, 'call-statement-forbids-parens')).toHaveLength(1);
	});

	it('uses nested With receivers from outer leading-dot expressions', () => {
		const person = 'Public Property Get Child() As Child\nEnd Property\n';
		const child =
			'Public Property Get Age() As Integer\nEnd Property\n' +
			'Public Sub Save(ByVal Count As Long)\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        With .Child\n' +
			'            .Delete\n' +
			'            .Age = 2\n' +
			'            .Save "bad"\n' +
			'        End With\n' +
			'    End With\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
				{ moduleName: 'Child', moduleKind: 'class', source: child },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'Delete' });
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(1);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(1);
	});

	it('uses parenthesized member receivers for source-backed class diagnostics', () => {
		const person = 'Public Property Get Child() As Child\nEnd Property\n';
		const child = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    (p.Child).Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
					{ moduleName: 'Child', moduleKind: 'class', source: child },
				]),
			}),
			'member-not-found',
		);

		expectDiagnostic(src, hits, 'member-not-found', { span: 'Delete' });
	});

	it('flags Exit For and Exit Do outside matching loops', () => {
		const src = 'Sub T()\n    Exit For\n    Exit Do\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-outside-block');
		expect(hits).toHaveLength(2);
		expect(spanText(src, hits[0])).toBe('Exit For');
		expect(spanText(src, hits[1])).toBe('Exit Do');
	});

	it('accepts Exit For and Exit Do inside matching loops', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n        Exit For\n    Next i\n' +
			'    Do\n        Exit Do\n    Loop\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-outside-block')).toHaveLength(0);
	});

	it('flags a Next variable that does not match the active For loop', () => {
		const src =
			'Sub T()\n' +
			'    Dim i As Long\n' +
			'    Dim j As Long\n' +
			'    For i = 1 To 3\n' +
			'        Debug.Print i\n' +
			'    Next j\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'next-variable-mismatch');

		expectDiagnostic(src, hits, 'next-variable-mismatch', { severity: 'error', span: 'j' });
	});

	it('flags a Next variable that does not match the active For Each loop', () => {
		const src =
			'Sub T()\n' +
			'    For Each item In items\n' +
			'    Next other\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'next-variable-mismatch');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('other');
	});

	it('accepts omitted, matching, and nested Next variables', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n' +
			'        For j = 1 To 3\n' +
			'        Next J\n' +
			'    Next i\n' +
			'    For k = 1 To 3\n' +
			'    Next\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'next-variable-mismatch')).toHaveLength(0);
	});

	it('does not report Next variable mismatches from inactive branches', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n' +
			'#If VBA7 Then\n' +
			'    Next j\n' +
			'#Else\n' +
			'    Next i\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'next-variable-mismatch',
			),
		).toHaveLength(0);
	});

	it('flags a For Each control variable declared as a scalar intrinsic type', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Long\n' +
			'    For Each item In Array(1, 2, 3)\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-control-variable-type');

		expectDiagnostic(src, hits, 'for-each-control-variable-type', {
			severity: 'error',
			span: 'item',
		});
	});

	it('flags an array variable used as a For Each control variable', () => {
		const src =
			'Sub T()\n' +
			'    Dim values() As Variant\n' +
			'    For Each values In Array(1, 2, 3)\n' +
			'        Debug.Print values\n' +
			'    Next values\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-control-variable-type');

		expectDiagnostic(src, hits, 'for-each-control-variable-type', { span: 'values' });
	});

	it('uses visible exported scalar and array globals for For Each control variables', () => {
		const caller =
			'Sub T()\n' +
			'    For Each SharedItem In Array(1, 2, 3)\n' +
			'    Next SharedItem\n' +
			'    For Each SharedValues In Array(1, 2, 3)\n' +
			'    Next SharedValues\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedItem As Long\n' +
					'Public SharedValues() As Variant\n',
			},
		], 'Caller');
		const hits = byCode(diagnostics, 'for-each-control-variable-type');

		expectDiagnostics(caller, hits, 'for-each-control-variable-type', [
			{ span: 'SharedItem', message: 'Long' },
			{ span: 'SharedValues', message: 'array variable' },
		]);
	});

	it('keeps local Variant shadows and ambiguous exported For Each controls quiet', () => {
		const shadowCaller =
			'Sub T()\n' +
			'    Dim SharedItem As Variant\n' +
			'    For Each SharedItem In Array(1, 2, 3)\n' +
			'    Next SharedItem\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedItem As Long\n' },
		], 'Caller'), 'for-each-control-variable-type')).toHaveLength(0);

		const ambiguousCaller =
			'Sub T()\n' +
			'    For Each SharedItem In Array(1, 2, 3)\n' +
			'    Next SharedItem\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedItem As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedItem As String\n' },
		], 'Caller'), 'for-each-control-variable-type')).toHaveLength(0);
	});

	it('accepts Variant, Object, and host object For Each control variables', () => {
		const src =
			'Sub T()\n' +
			'    Dim value As Variant\n' +
			'    For Each value In Array(1, 2, 3)\n' +
			'    Next value\n' +
			'    Dim obj As Object\n' +
			'    For Each obj In items\n' +
			'    Next obj\n' +
			'    Dim ws As Worksheet\n' +
			'    For Each ws In ThisWorkbook.Worksheets\n' +
			'    Next ws\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'for-each-control-variable-type')).toHaveLength(0);
	});

	it('flags a project UDT used as a For Each control variable when type metadata is available', () => {
		const modules: ProjectTestModule[] = [
			{
				moduleName: 'Types',
				moduleKind: 'standard',
				source: 'Public Type TItem\n    Value As Long\nEnd Type\n',
			},
			{ moduleName: 'Module1', moduleKind: 'standard', source: '' },
		];
		const src =
			'Sub T()\n' +
			'    Dim item As TItem\n' +
			'    For Each item In items\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(src, modules, 'Module1'),
			'for-each-control-variable-type',
		);

		expectDiagnostic(src, hits, 'for-each-control-variable-type', { span: 'item' });
	});

	it('does not report For Each control variable type diagnostics from inactive branches', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Long\n' +
			'#If VBA7 Then\n' +
			'    For Each item In Array(1, 2, 3)\n' +
			'    Next item\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'for-each-control-variable-type',
			),
		).toHaveLength(0);
	});

	it('flags a For Each source declared as a scalar intrinsic type', () => {
		const src =
			'Private ModuleValue As String\n' +
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim Value As Long\n' +
			'    For Each item In Value\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'    For Each item In ModuleValue\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-source-type');

		expectDiagnostics(src, hits, 'for-each-source-type', [
			{ severity: 'error', span: 'Value' },
			{ span: 'ModuleValue' },
		]);
	});

	it('uses visible exported scalar globals for For Each source types', () => {
		const caller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
			], 'Caller'),
			'for-each-source-type',
		);

		expectDiagnostic(caller, hits, 'for-each-source-type', { span: 'SharedValue' });
	});

	it('keeps exported arrays Variants local shadows and ambiguous For Each sources quiet', () => {
		const acceptedCaller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValues\n' +
			'    Next item\n' +
			'    For Each item In SharedFlexible\n' +
			'    Next item\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(acceptedCaller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValues() As Long\n' +
					'Public SharedFlexible As Variant\n',
			},
		], 'Caller'), 'for-each-source-type')).toHaveLength(0);

		const shadowCaller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim SharedValue() As Long\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
		], 'Caller'), 'for-each-source-type')).toHaveLength(0);

		const ambiguousCaller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue As String\n' },
		], 'Caller'), 'for-each-source-type')).toHaveLength(0);
	});

	it('accepts known array/object-like and unresolved For Each sources', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim values() As Long\n' +
			'    Dim maybeValues As Variant\n' +
			'    Dim obj As Object\n' +
			'    For Each item In values\n' +
			'    Next item\n' +
			'    For Each item In maybeValues\n' +
			'    Next item\n' +
			'    For Each item In obj\n' +
			'    Next item\n' +
			'    For Each item In obj.Items\n' +
			'    Next item\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'for-each-source-type')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking For Each source types', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    For Each item In Values\n' +
			'    Next item\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'for-each-source-type',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'for-each-source-type',
			),
		).toHaveLength(0);
	});

	it('does not report For Each source diagnostics from inactive branches', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim Value As Long\n' +
			'#If VBA7 Then\n' +
			'    For Each item In Value\n' +
			'    Next item\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'for-each-source-type',
			),
		).toHaveLength(0);
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

// Exact diagnostic wording is pinned here once per rule; behavior tests above
// assert rule code + severity + span via expectDiagnostic and stay wording-free.
describe('diagnostic message wording', () => {
	it('pins the message for ambiguous-enum-member', () => {
		const src =
			'Public Enum ENeg_AmbiguousOne\n' +
			'    NegAmbiguousValue = 1\n' +
			'End Enum\n' +
			'\n' +
			'Public Enum ENeg_AmbiguousTwo\n' +
			'    NegAmbiguousValue = 2\n' +
			'End Enum\n' +
			'\n' +
			'Public Sub T()\n' +
			'    Debug.Print NegAmbiguousValue\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'ambiguous-enum-member');
		expect(hits[0].message).toBe("Ambiguous Enum member reference: 'NegAmbiguousValue' is defined by multiple visible Enums (ENeg_AmbiguousOne, ENeg_AmbiguousTwo). Qualify the reference with an Enum or module name.");
	});

	it('pins the message for argument-count', () => {
		const src = 'Sub Main()\n    MsgBox()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits[0].message).toBe("Wrong number of arguments to 'MsgBox': expected between 1 and 5 arguments, but got 0.");
	});

	it('pins the message for argument-object-type-mismatch', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Left$("abcdef", 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits[0].message).toBe("Argument 'item' of 'NeedsObject' expects Object, but got Left$(...) As String. An object parameter requires an object value.");
	});

	it('pins the message for argument-type-mismatch', () => {
		const src = 'Sub T()\n    x = Left("abcdef", "bad")\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits[0].message).toBe("Argument 'Length' of 'Left' expects Long, but got String literal \"bad\". This string literal cannot be converted to a numeric value. This will raise Run-time error '13': Type mismatch.");
	});

	it('pins the message for array-assignment-to-scalar', () => {
		const src =
			'Private ModuleValues(1 To 3) As Long\n' +
			'Public Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    Dim DynamicValues() As String\n' +
			'    Dim Value As Long\n' +
			'    Dim ScalarText As String\n' +
			'    Value = Values\n' +
			'    ScalarText = DynamicValues\n' +
			'    Value = ModuleValues\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'array-assignment-to-scalar');
		expect(hits[0].message).toBe("Array variable 'Values' cannot be assigned to scalar 'Value'. Assign an array element or use a Variant/array target.");
	});

	it('pins the message for array-bound-requires-array', () => {
		const caller =
			'Public Sub T()\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
			], 'Caller'),
			'array-bound-requires-array',
		);
		expect(hits[0].message).toBe("LBound requires an array argument, but 'SharedValue' is declared As Long.");
	});

	it('pins the message for assignment-object-type-mismatch', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim wb As Workbook\n' +
			'    Dim rng As Range\n' +
			'    Set rng = ActiveSheet.Range("A1")\n' +
			'    Set wb = ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-object-type-mismatch');
		expect(hits[0].message).toBe("Object assignment to 'wb' expects Workbook, but got ActiveSheet.Range(\"A1\") As Excel.Range. This object type is not compatible with Workbook.");
	});

	it('pins the message for assignment-type-mismatch', () => {
		const src =
			'Public Function Total() As Double\n' +
			'    Total = "blah"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');
		expect(hits[0].message).toBe("Assignment to 'Total' expects Double, but got String literal \"blah\". This string literal cannot be converted to a numeric value. This will raise Run-time error '13': Type mismatch.");
	});

	it('pins the message for byref-argument-type-mismatch', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'byref-argument-type-mismatch');
		expect(hits[0].message).toBe("ByRef argument 'value' of 'Mutate' expects Long, but 'amount' is declared as Integer. This is a VBE compile error: ByRef argument type mismatch.");
	});

	it('pins the message for call-statement-forbids-parens', () => {
		const src = 'Sub T()\n    DoEvents()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits[0].message).toBe("Standalone 'DoEvents()' cannot use empty parentheses in statement context; use 'DoEvents' as a statement or use it in an expression.");
	});

	it('pins the message for division-by-zero', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'    a = 1 \\ 0\n' +
			'    a = 1 Mod 0\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits[0].message).toBe("Expression uses '/' with a zero divisor. This will raise Run-time error '11': Division by zero.");
	});

	it('pins the message for duplicate-enum-member', () => {
		const src =
			'Public Enum ENeg_DuplicateMembers\n' +
			'    NegEnumShared = 1\n' +
			'    NegEnumShared = 2\n' +
			'End Enum\n';
		const hits = byCode(analyzeModule(src), 'duplicate-enum-member');
		expect(hits[0].message).toBe("Duplicate Enum member 'NegEnumShared' in Enum 'ENeg_DuplicateMembers'.");
	});

	it('pins the message for duplicate-label', () => {
		const src =
			'Sub T()\n' +
			'StartHere:\n' +
			'    Debug.Print "first"\n' +
			'StartHere:\n' +
			'    Debug.Print "second"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-label');
		expect(hits[0].message).toBe("Label 'StartHere' is already defined in procedure 'T'.");
	});

	it('pins the message for else-branch-order', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#Else\n' +
			'    Debug.Print 1\n' +
			'#ElseIf True Then\n' +
			'    Debug.Print 2\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');
		expect(hits[0].message).toBe("'#ElseIf' cannot appear after '#Else' in the same conditional-compilation block.");
	});

	it('pins the message for erase-requires-array', () => {
		const src =
			'Private ModuleValue As Long\n' +
			'Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim Value As Long\n' +
			'    Erase obj, Value, ModuleValue\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'erase-requires-array');
		expect(hits[0].message).toBe("Erase target 'obj' must be an array or Variant, but it is declared As Object.");
	});

	it('pins the message for event-declaration-module-kind', () => {
		const src =
			'Public Event BeforeAdd(ByRef arr As Variant, ByRef cancel As Boolean)\n' +
			'Private Event AfterAdd(ByRef arr As Variant)\n';
		const hits = byCode(analyzeModule(src), 'event-declaration-module-kind');
		expect(hits[0].message).toBe("Event declaration 'BeforeAdd' is only valid in class, document, or UserForm modules.");
	});

	it('pins the message for event-handler-module-scope', () => {
		const src = 'Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Module1', moduleKind: 'standard' }),
			'event-handler-module-scope',
		);
		expect(hits[0].message).toBe("'Workbook_Open' matches a Workbook event handler, but this standard module is not where Excel wires that event. It will behave like an ordinary procedure here.");
	});

	it('pins the message for fixed-array-redim', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    ReDim Values(1 To 10) As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-array-redim');
		expect(hits[0].message).toBe("Fixed-size array 'Values' cannot be resized with ReDim.");
	});

	it('pins the message for fixed-length-string-size', () => {
		const src =
			'Private Const HeaderCodeLength As Long = &H14\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const LocalTooSmall As Long = &O0\n' +
			'    Dim localName As String * LocalTooSmall\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits[0].message).toBe("Fixed-length String size must be between 1 and 65526 characters; got 0.");
	});

	it('pins the message for for-each-control-variable-type', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Long\n' +
			'    For Each item In Array(1, 2, 3)\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-control-variable-type');
		expect(hits[0].message).toBe("For Each control variable 'item' must be Variant or Object, but it is declared As Long.");
	});

	it('pins the message for for-each-source-type', () => {
		const caller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
			], 'Caller'),
			'for-each-source-type',
		);
		expect(hits[0].message).toBe("For Each source 'SharedValue' must be a collection object or array, but it is declared As Long.");
	});

	it('pins the message for friend-declaration', () => {
		const src = 'Friend Sub InternalOnly()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'friend-declaration');
		expect(hits[0].message).toBe("Friend procedure 'InternalOnly' is only valid in class, document, or UserForm modules.");
	});

	it('pins the message for implements-statement-placement', () => {
		const src = 'Option Explicit\nImplements Person\n';
		const hits = byCode(analyzeModule(src), 'implements-statement-placement');
		expect(hits[0].message).toBe("Implements statement 'Person' is only valid in class, document, or UserForm modules.");
	});

	it('pins the message for invalid-as-type-name', () => {
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
		expect(hits[0].message).toBe("'Status' is ambiguous because multiple visible project types use that name.");
	});

	it('pins the message for invalid-erase-target', () => {
		const src =
			'Sub T()\n' +
			'    Erase 1 + 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-erase-target');
		expect(hits[0].message).toBe("Erase target must be a variable or array name, not an arbitrary expression.");
	});

	it('pins the message for invalid-expression-syntax', () => {
		const src = 'Sub T()\n    value = flag ? 1 : 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits[0].message).toBe("VBA does not support the '?' conditional operator in code modules; use If...Then...Else, or IIf(...) only when both branches are safe to evaluate.");
	});

	it('pins the message for member-not-found', () => {
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.AfterSave True\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits[0].message).toBe("Method or data member not found: 'Excel.Workbook.AfterSave'.");
	});

	it('pins the message for missing-return-assignment', () => {
		const src =
			'Public Function myFunction()\n' +
			'\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');
		expect(hits[0].message).toBe("Function 'myFunction' has no return assignment; VBA will return the default value. Assign to 'myFunction' before exit if a value is intended.");
	});

	it('pins the message for module-declaration-after-procedure', () => {
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
		expect(hits[0].message).toBe("Declare statements in the active conditional-compilation branch belong in the module declarations section, before procedures.");
	});

	it('pins the message for module-declaration-in-procedure', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "body"\n' +
			'    Attribute T.VB_Description = "bad placement"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-in-procedure');
		expect(hits[0].message).toBe("Attribute statements must appear in the module declarations section, not inside a procedure.");
	});

	it('pins the message for next-variable-mismatch', () => {
		const src =
			'Sub T()\n' +
			'    Dim i As Long\n' +
			'    Dim j As Long\n' +
			'    For i = 1 To 3\n' +
			'        Debug.Print i\n' +
			'    Next j\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'next-variable-mismatch');
		expect(hits[0].message).toBe("Next variable 'j' does not match active For control variable 'i'.");
	});

	it('pins the message for non-callable-call', () => {
		const src = 'Sub Main()\n    Dim testStr As String\n    testStr\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits[0].message).toBe("Cannot call 'testStr' because it resolves to a local variable, not a Sub or Function.");
	});

	it('pins the message for object-module-public-member', () => {
		const src =
			'Public Const MaxRows As Long = 1000\n' +
			'Public Names() As String\n' +
			'Public FixedName As String * 20\n' +
			'Public Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
			'object-module-public-member',
		);
		expect(hits[0].message).toBe("Public constants are not allowed as Public members of object modules; VBE Compile rejects this declaration.");
	});

	it('pins the message for object-variable-not-set', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    obj.ToString\n' +
			'    Dim ws As Worksheet\n' +
			'    ws.Range("A1").Value = 1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'object-variable-not-set');
		expect(hits[0].message).toBe("Object variable 'obj' is Nothing before member access. This will raise Run-time error '91': Object variable or With block variable not set.");
	});

	it('pins the message for paramarray-non-variant', () => {
		const src =
			'Public Sub Combined012TypedParamarray(ParamArray values() As String)\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-non-variant');
		expect(hits[0].message).toBe("ParamArray 'values' elements must be Variant, but this parameter is declared As String.");
	});

	it('pins the message for paramarray-with-optional', () => {
		const src =
			'Public Sub Combined011ParamarrayWithOptional(Optional ByVal prefix As String = "x", ParamArray values() As Variant)\n' +
			'    Debug.Print prefix\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-with-optional');
		expect(hits[0].message).toBe("ParamArray 'values' cannot be used in the same parameter list as Optional arguments.");
	});

	it('pins the message for parameter-array-as-type-syntax', () => {
		const src =
			'Public Sub NegParam06_BadArrayParameterSyntax(ByVal values As Long())\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, 'parameter-array-as-type-syntax');
		expect(hits[0].message).toBe("Array parameter 'values' must place parentheses after the parameter name, before the As clause; use 'values() As Long'.");
	});

	it('pins the message for parameter-default-type-mismatch', () => {
		const src =
			'Public Sub NegParam07_OptionalObjectDefaultNonNothing(Optional ByVal obj As Object = 1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'parameter-default-type-mismatch');
		expect(hits[0].message).toBe("Optional parameter 'obj' expects Object, but its default value is numeric literal 1. Optional object parameter defaults must be Nothing.");
	});

	it('pins the message for property-accessor-signature-mismatch', () => {
		const src =
			'Public Property Get Item(ByVal index As Long) As String\n' +
			'End Property\n' +
			'Public Property Let Item(ByVal value As String)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-accessor-signature-mismatch');
		expect(hits[0].message).toBe("Property Let 'Item' argument list must match Property Get 'Item' before the final value parameter. Expected 1 index parameter, but found 0.");
	});

	it('pins the message for property-let-object-value', () => {
		const src =
			'Public Property Let NegProp06_LetObjectValue(ByVal Value As Object)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-let-object-value');
		expect(hits[0].message).toBe("Property Let 'NegProp06_LetObjectValue' final value parameter 'Value' must not be an object reference; use Property Set because it is declared As Object.");
	});

	it('pins the message for property-set-scalar-value', () => {
		const src =
			'Public Property Set Number(ByVal value As Long)\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-set-scalar-value');
		expect(hits[0].message).toBe("Property Set 'Number' final value parameter 'value' must be an object reference, but it is declared As Long.");
	});

	it('pins the message for property-setter-missing-value', () => {
		const src =
			'Public Property Let Name()\n' +
			'End Property\n';
		const hits = byCode(analyzeModule(src), 'property-setter-missing-value');
		expect(hits[0].message).toBe("Property Let 'Name' must include a final value parameter.");
	});

	it('pins the message for raiseevent-undeclared-event', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub Touch()\n' +
			'    RaiseEvent Changed()\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			moduleName: 'EventSource',
			moduleKind: 'class',
			knownIdentifiers: new Set<string>(),
		});
		expect(diagnostics[0].message).toBe("Event 'Changed' is not declared in this module, so it cannot be raised with RaiseEvent.");
	});

	it('pins the message for readonly-member-assignment', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'readonly-member-assignment',
		);
		expect(hits[0].message).toBe("Cannot assign to read-only property 'p.Age'.");
	});

	it('pins the message for redim-preserve-dimension-change', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 3, 1 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');
		expect(hits[0].message).toBe("ReDim Preserve can only resize the last dimension of 'grid'. Dimension 1 changes before the final dimension.");
	});

	it('pins the message for runtime-argument-value', () => {
		const src =
			'Public Function Left(ByVal text As String, ByVal count As Long) As String\n' +
			'End Function\n' +
			'Sub T()\n' +
			'    Dim localValue As Long\n' +
			'    Dim Left As Long\n' +
			'    a = Left("abcdef", -1)\n' +
			'    b = VBA.Left$("abcdef", -1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');
		expect(hits[0].message).toBe("Argument 'Length' of 'Left$' is -1; this will raise Run-time error '5': Invalid procedure call or argument.");
	});

	it('pins the message for runtime-conversion-value', () => {
		const src =
			'Sub T()\n' +
			'    Dim CDate As Variant\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("not a date")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-conversion-value');
		expect(hits[0].message).toBe("VBA.CDate cannot convert \"not a date\" to Date. This will raise Run-time error '13': Type mismatch.");
	});

	it('pins the message for scalar-member-access', () => {
		const src = 'Sub Main()\n    Dim value As String\n    value.\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits[0].message).toBe("Member access on 'value' is invalid because it is declared as String. This is a VBE compile error: Syntax error.");
	});

	it('pins the message for scalar-redim', () => {
		const src =
			'Private ModuleValue As Long\n' +
			'Sub T()\n' +
			'    Dim Value As Long\n' +
			'    ReDim Value(1 To 10)\n' +
			'    ReDim Preserve ModuleValue(1 To 10)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-redim');
		expect(hits[0].message).toBe("Scalar variable 'Value' cannot be resized with ReDim; declare it as a dynamic array first.");
	});

	it('pins the message for set-required', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-required');
		expect(hits[0].message).toBe("Object assignment to 'ws' requires Set because it is declared as Worksheet.");
	});

	it('pins the message for set-requires-object', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Text As String\n' +
			'    Set text = New Collection\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-requires-object');
		expect(hits[0].message).toBe("Set assignment requires an object variable, but 'text' is declared as String.");
	});

	it('pins the message for statement-outside-procedure', () => {
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
		expect(hits[0].message).toBe("NegAmbiguousValue statement is invalid outside a Sub, Function, or Property procedure.");
	});

	it('pins the message for string-arithmetic-coercion', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As Integer\n' +
			'    shouldErrorTest1 = 1 + "string"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'string-arithmetic-coercion');
		expect(hits[0].message).toBe("Assignment to 'shouldErrorTest1' expects Integer, but this numeric expression contains nonnumeric string literal \"string\". This will raise Run-time error '13': Type mismatch.");
	});

	it('pins the message for type-declaration-character-as-clause', () => {
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
		expect(hits[0].message).toBe("Type field 'Code' combines type-declaration character '$' with an As clause; use only one type declaration form.");
	});

	it('pins the message for unallocated-dynamic-array-access', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Debug.Print values(0)\n' +
			'    values(1) = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');
		expect(hits[0].message).toBe("Dynamic array 'values' is not allocated before indexed access. This will raise Run-time error '9': Subscript out of range.");
	});

	it('pins the message for undeclared-variable', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim declared As Long\n' +
			'    Dim obj As Object\n' +
			'10 declared = 1\n' +
			'20 Set obj = ActiveSheet\n' +
			'30 missing = 1\n' +
			'40 Set missingObj = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits[0].message).toBe("Variable not defined: 'missing'. Declare it before assigning to it, or remove Option Explicit.");
	});

	it('pins the message for unexpected-declaration-token', () => {
		const src = 'Sub T()\n    Dim s1 As String thisshoulderror\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unexpected-declaration-token');
		expect(hits[0].message).toBe("Unexpected token 'thisshoulderror' after a complete declaration type; this will fail to compile as a syntax error.");
	});

	it('pins the message for withevents-declaration', () => {
		const src = 'Private WithEvents App As Application\n';
		const hits = byCode(analyzeModule(src), 'withevents-declaration');
		expect(hits[0].message).toBe("WithEvents variable 'App' is only valid in class, document, or UserForm modules.");
	});
});
