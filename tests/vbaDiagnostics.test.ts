import { describe, it, expect } from 'vitest';
import {
	analyzeModule,
	DIAGNOSTIC_RULES,
	STRUCTURAL_DIAGNOSTIC_RULES,
	diagnosticMetadataForCode,
	diagnosticSourceForCode,
	isXlideDiagnosticSource,
} from '../src/analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
} from '../src/vbaProjectAnalysis';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from './helpers/diagnostics';
import {
	analyzeProjectModule,
	projectClassMembers,
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
