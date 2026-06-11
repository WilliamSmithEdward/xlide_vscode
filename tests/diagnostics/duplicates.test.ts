// Diagnostics tests: duplicates rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

describe('analyzeModule - duplicate procedures', () => {
	it('flags two Subs with the same name', () => {
		const src = 'Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-procedure');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Foo');
	});

	it('flags a Function colliding with a Sub of the same name', () => {
		const src = 'Sub Foo()\nEnd Sub\nFunction Foo()\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(1);
	});

	it('allows Property Get/Let/Set to share a name', () => {
		const src =
			'Property Get Item() As Long\nEnd Property\n' +
			'Property Let Item(v As Long)\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
	});

	it('flags a duplicate Property Get', () => {
		const src =
			'Property Get Item() As Long\nEnd Property\n' +
			'Property Get Item() As Long\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(1);
	});

	it('filters duplicate procedures in default VBA7 conditional branches', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#Else\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#End If\n';

		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-procedure',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate declarations in scope', () => {
	it('flags the same local declared twice', () => {
		const src = 'Sub T()\n    Dim x As Long\n    Dim x As String\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-declaration');
		expect(hits).toHaveLength(1);
	});

	it('flags a local colliding with a parameter', () => {
		const src = 'Sub T(ByVal n As Long)\n    Dim n As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('flags duplicate parameter names', () => {
		const src = 'Sub T(a As Long, a As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('treats procedure scope as flat across branches', () => {
		const src =
			'Sub T()\n' +
			'    If True Then\n        Dim x As Long\n    End If\n' +
			'    Dim x As Long\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('filters duplicate declarations in default VBA7 conditional branches', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim value As LongPtr\n' +
			'#Else\n' +
			'    Dim value As Long\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'duplicate-declaration',
			),
		).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-declaration',
			),
		).toHaveLength(0);
	});

	it('does not flag distinct local names', () => {
		const src = 'Sub T()\n    Dim x As Long\n    Dim y As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate module members', () => {
	it('flags the same module variable declared twice', () => {
		const src = 'Private Total As Long\nPublic Total As String\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(1);
	});

	it('filters duplicate module variables in default VBA7 conditional branches', () => {
		const src =
			'#If VBA7 Then\n' +
			'Private activeSheetPtr As LongPtr\n' +
			'#Else\n' +
			'Private activeSheetPtr As Long\n' +
			'#End If\n';

		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-module-variable',
			),
		).toHaveLength(0);
	});

	it('filters duplicate module variables in default Win64 conditional branches', () => {
		const src =
			'#If Win64 Then\n' +
			'Private nativeHandle As LongPtr\n' +
			'#Else\n' +
			'Private nativeHandle As Long\n' +
			'#End If\n';

		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { Win64: false } },
				}),
				'duplicate-module-variable',
			),
		).toHaveLength(0);
	});

	it('does not flag distinct module variables', () => {
		const src = 'Private A As Long\nPrivate B As Long\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate Enum members', () => {
	it('flags duplicate member names inside the same Enum block', () => {
		const src =
			'Public Enum ENeg_DuplicateMembers\n' +
			'    NegEnumShared = 1\n' +
			'    NegEnumShared = 2\n' +
			'End Enum\n';
		const hits = byCode(analyzeModule(src), 'duplicate-enum-member');

		expectDiagnostic(src, hits, 'duplicate-enum-member', {
			severity: 'error',
			span: 'NegEnumShared',
		});
	});

	it('treats Enum member duplicate checks as case-insensitive', () => {
		const src =
			'Private Enum Mode\n' +
			'    Ready = 1\n' +
			'    READY = 2\n' +
			'End Enum\n';
		const hits = byCode(analyzeModule(src), 'duplicate-enum-member');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('READY');
	});

	it('does not report duplicates from inactive whole Enum branches', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Enum ConditionalMode\n' +
			'    SharedMode = 1\n' +
			'    SharedMode = 2\n' +
			'End Enum\n' +
			'#End If\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-enum-member',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - ambiguous Enum member references', () => {
	it('does not flag same-name members declared in separate Enum blocks by themselves', () => {
		const src =
			'Public Enum ENeg_AmbiguousOne\n' +
			'    NegAmbiguousValue = 1\n' +
			'End Enum\n' +
			'\n' +
			'Public Enum ENeg_AmbiguousTwo\n' +
			'    NegAmbiguousValue = 2\n' +
			'End Enum\n';

		expect(byCode(analyzeModule(src), 'ambiguous-enum-member')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(0);
	});

	it('flags an unqualified read of a member shared by separate same-module Enums', () => {
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

		expectDiagnostic(src, hits, 'ambiguous-enum-member', {
			severity: 'error',
			span: 'NegAmbiguousValue',
		});
	});

	it('does not flag qualified reads or local shadows of ambiguous Enum member names', () => {
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
			'    Dim NegAmbiguousValue As Long\n' +
			'    Debug.Print NegAmbiguousValue\n' +
			'    Debug.Print ENeg_AmbiguousOne.NegAmbiguousValue\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'ambiguous-enum-member')).toHaveLength(0);
	});

	it('flags externally ambiguous Enum member reads unless the current module binds the name', () => {
		const caller =
			'Public Sub T()\n' +
			'    Debug.Print SharedModeValue\n' +
			'End Sub\n';
		const globalsA =
			'Public Enum SharedModeA\n' +
			'    SharedModeValue = 1\n' +
			'End Enum\n';
		const globalsB =
			'Public Enum SharedModeB\n' +
			'    SharedModeValue = 2\n' +
			'End Enum\n';
		const modules = [
			{ moduleName: 'Caller', source: caller },
			{ moduleName: 'GlobalsA', source: globalsA },
			{ moduleName: 'GlobalsB', source: globalsB },
		];
		const hits = byCode(
			analyzeProjectModule(caller, modules, 'Caller'),
			'ambiguous-enum-member',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedModeValue');

		const localCaller =
			'Public Enum LocalMode\n' +
			'    SharedModeValue = 0\n' +
			'End Enum\n' +
			'\n' +
			caller;
		expect(
			byCode(
				analyzeProjectModule(localCaller, modules, 'Caller'),
				'ambiguous-enum-member',
			),
		).toHaveLength(0);

		const moduleShadowCaller =
			'Private SharedModeValue As Long\n' +
			'\n' +
			caller;
		expect(
			byCode(
				analyzeProjectModule(moduleShadowCaller, modules, 'Caller'),
				'ambiguous-enum-member',
			),
		).toHaveLength(0);
	});
});
