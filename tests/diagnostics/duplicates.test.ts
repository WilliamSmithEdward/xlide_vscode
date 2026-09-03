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

	it('allows one procedure per arm when the condition is not decidable', () => {
		const src =
			'#If HostIsExcel Then\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#ElseIf HostIsWord Then\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#Else\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
	});

	it('allows one Property Get per arm when the condition is not decidable', () => {
		const src =
			'#If HostIsExcel Then\n' +
			'Property Get Item() As Long\nItem = 1\nEnd Property\n' +
			'#Else\n' +
			'Property Get Item() As Long\nItem = 2\nEnd Property\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
	});

	it('still flags two procedures in one arm', () => {
		const src =
			'#If HostIsExcel Then\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#End If\n';
		const hits = byCode(analyzeModule(src), 'duplicate-procedure');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Configure');
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

	it('allows one local per arm when the condition is not decidable', () => {
		const src =
			'Sub T()\n' +
			'#If Whatever Then\n    Dim value As Variant\n' +
			'#ElseIf AnotherExample Then\n    Dim value As String\n' +
			'#Else\n    Dim value As Long\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(0);
	});

	it('still flags a local repeated within one arm', () => {
		const src =
			'Sub T()\n' +
			'#If Whatever Then\n    Dim value As Long\n    Dim value As String\n#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('still flags a local colliding with a parameter from inside an arm', () => {
		const src =
			'Sub T(ByVal value As Long)\n' +
			'#If Whatever Then\n    Dim value As String\n#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
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

	// github.com/WilliamSmithEdward/xlide_vscode/issues/58: the arms of one
	// `#If` chain are alternatives, so a name declared in each is declared once
	// in whichever build the compiler makes. That holds however the conditions
	// evaluate; a decidable condition already leaves the losing arms inactive,
	// so the arms XLIDE cannot decide are where it matters.
	it('allows one module variable per arm when the condition is not decidable', () => {
		const src =
			'#If Whatever Then\n    Dim Test As Variant\n' +
			'#ElseIf AnotherExample Then\n    Dim Test As String\n' +
			'#Else\n    Dim test As Long\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});

	it('allows one Const per arm when the condition is not decidable', () => {
		const src =
			'#If Whatever Then\nPrivate Const Limit As Long = 1\n' +
			'#Else\nPrivate Const Limit As Long = 2\n#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});

	it('allows a name in an outer arm and in a chain nested in the other arm', () => {
		const src =
			'#If Whatever Then\nDim Test As Long\n' +
			'#Else\n#If AnotherExample Then\nDim Test As String\n#End If\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});

	it('still flags a name repeated within one arm', () => {
		const src = '#If Whatever Then\nDim Test As Long\nDim Test As String\n#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(1);
	});

	it('still flags an unconditional name repeated inside an arm', () => {
		const src = 'Dim Test As Long\n#If Whatever Then\nDim Test As String\n#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(1);
	});

	it('still flags a name shared by two separate #If chains, which can both compile', () => {
		const src =
			'#If Whatever Then\nDim Test As Long\n#End If\n' +
			'#If AnotherExample Then\nDim Test As String\n#End If\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(1);
	});

	it('still flags a name repeated in one arm alongside a legitimate use of the other', () => {
		const src =
			'#If Whatever Then\nDim Test As Long\nDim Test As String\n' +
			'#Else\nDim Test As Double\n#End If\n';
		const hits = byCode(analyzeModule(src), 'duplicate-module-variable');
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBeLessThan(src.indexOf('#Else'));
	});

	it('reports repeats in source order, not grouped by name', () => {
		const src =
			'Dim Alpha As Long\nDim Beta As Long\n' +
			'Dim Beta As String\nDim Alpha As String\n';
		const hits = byCode(analyzeModule(src), 'duplicate-module-variable');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Beta', 'Alpha']);
	});

	it('does not flag distinct module variables', () => {
		const src = 'Private A As Long\nPrivate B As Long\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate Enum members', () => {
	// These two rules used to skip any member in a branch they could not
	// decide, which kept them quiet across arms (right) and also inside a
	// single arm (wrong). They now drop only provably-inactive members and ask
	// the arm question for the rest (issues/58).
	it('flags an Enum member repeated inside one arm of an undecidable chain', () => {
		const src =
			'Public Enum E\n#If CUSTOM_FLAG Then\n    A = 1\n    A = 2\n#End If\nEnd Enum\n';
		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(1);
	});

	it('allows an Enum member once per arm', () => {
		const src =
			'Public Enum E\n#If CUSTOM_FLAG Then\n    A = 1\n#Else\n    A = 2\n#End If\nEnd Enum\n';
		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(0);
	});

	it('ignores a repeat inside a provably dead arm', () => {
		const src =
			'Public Enum E\n#If Mac Then\n    A = 1\n    A = 2\n#End If\n    B = 3\nEnd Enum\n';
		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(0);
	});

	it('flags a Type field repeated inside one arm of an undecidable chain', () => {
		const src =
			'Public Type P\n#If CUSTOM_FLAG Then\n    F As Long\n    F As String\n#End If\nEnd Type\n';
		expect(byCode(analyzeModule(src), 'duplicate-type-field')).toHaveLength(1);
	});

	it('allows a Type field once per arm', () => {
		const src =
			'Public Type P\n#If CUSTOM_FLAG Then\n    F As Long\n#Else\n    F As String\n#End If\nEnd Type\n';
		expect(byCode(analyzeModule(src), 'duplicate-type-field')).toHaveLength(0);
	});

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

	it('does not report a duplicate across an #If branch inside the Enum body', () => {
		// The two SharedMode definitions live in mutually exclusive branches, so
		// only one is ever compiled - not a duplicate.
		const src =
			'Public Enum ConditionalMode\n' +
			'#If 0 Then\n' +
			'    SharedMode = 1\n' +
			'#Else\n' +
			'    SharedMode = 2\n' +
			'#End If\n' +
			'End Enum\n';

		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(0);
	});

	it('does not report a duplicate across #If/#Else arms guarded by an unknown constant', () => {
		// HOST_BUILD is undefined, so both arms are "unknown" rather than provably
		// inactive - but #If and its #Else are mutually exclusive, so only one
		// member is ever compiled. Flagging a duplicate here would be a false
		// positive (the idiomatic per-platform member-value pattern).
		const src =
			'Public Enum PtrSize\n' +
			'#If HOST_BUILD Then\n' +
			'    Size = 8\n' +
			'#Else\n' +
			'    Size = 4\n' +
			'#End If\n' +
			'End Enum\n';

		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(0);
	});

	it('still flags two same-named members that are both provably active', () => {
		// One copy in the always-on region, one in a dead #If 0 branch (skipped),
		// and a third active copy: the two active copies collide.
		const src =
			'Public Enum Modes\n' +
			'    Dup = 1\n' +
			'#If 0 Then\n' +
			'    Dup = 2\n' +
			'#End If\n' +
			'    Dup = 3\n' +
			'End Enum\n';

		expect(byCode(analyzeModule(src), 'duplicate-enum-member')).toHaveLength(1);
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

describe('analyzeModule - duplicate Type fields', () => {
	it('flags duplicate field names inside the same Type block', () => {
		const src = 'Public Type TRec\n    Id As Long\n    Id As String\nEnd Type\n';
		const hits = byCode(analyzeModule(src), 'duplicate-type-field');

		expectDiagnostic(src, hits, 'duplicate-type-field', {
			severity: 'error',
			span: 'Id',
		});
	});

	it('treats Type field duplicate checks as case-insensitive', () => {
		const src = 'Private Type TRec\n    Name As String\n    NAME As String\nEnd Type\n';
		const hits = byCode(analyzeModule(src), 'duplicate-type-field');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('NAME');
	});

	it('stays quiet for a Type with all-distinct field names', () => {
		const src = 'Public Type TPoint\n    X As Long\n    Y As Long\n    Z As Long\nEnd Type\n';
		expect(byCode(analyzeModule(src), 'duplicate-type-field')).toHaveLength(0);
	});

	it('detects duplicates by the bare field name for array fields', () => {
		// The collision key is the field name, not the "(bounds)" suffix.
		const src = 'Public Type TRec\n    Items(10) As Long\n    Items As Long\nEnd Type\n';
		const hits = byCode(analyzeModule(src), 'duplicate-type-field');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Items');
	});

	it('does not report a duplicate across #If/#Else arms inside the Type body', () => {
		// Mutually exclusive arms guarded by an unknown constant: only one field is
		// ever compiled, so this is not a duplicate.
		const src =
			'Public Type TRec\n' +
			'#If HOST_BUILD Then\n' +
			'    Handle As LongPtr\n' +
			'#Else\n' +
			'    Handle As Long\n' +
			'#End If\n' +
			'End Type\n';
		expect(byCode(analyzeModule(src), 'duplicate-type-field')).toHaveLength(0);
	});

	it('still flags two same-named fields that are both provably active', () => {
		const src =
			'Public Type TRec\n' +
			'    Dup As Long\n' +
			'#If 0 Then\n' +
			'    Dup As String\n' +
			'#End If\n' +
			'    Dup As Double\n' +
			'End Type\n';
		expect(byCode(analyzeModule(src), 'duplicate-type-field')).toHaveLength(1);
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
