// Diagnostics tests: moduleKind rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';

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
