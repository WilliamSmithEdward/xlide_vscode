// Diagnostics tests: rules promoted from the syntax corpus after Excel/VBE oracle
// verification (2026-06-13) — empty-type, duplicate-option, duplicate-case-else.
// Each is vbe-oracle-verified; the matching oracle case is asserted in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

describe('analyzeModule - empty-type (PCEC_004)', () => {
	const CODE = 'empty-type';

	it('flags a user-defined Type with no members', () => {
		const src = 'Public Type EmptyType\nEnd Type\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('EmptyType');
	});

	it('flags a Private empty Type too', () => {
		expect(byCode(analyzeModule('Private Type T\nEnd Type\n'), CODE)).toHaveLength(1);
	});

	it('stays quiet for a Type with at least one member', () => {
		expect(byCode(analyzeModule('Public Type T\n    X As Long\nEnd Type\n'), CODE)).toHaveLength(0);
	});

	it('stays quiet for an empty Enum (VBE accepts it — PCEC_003 refuted)', () => {
		expect(byCode(analyzeModule('Public Enum E\nEnd Enum\n'), CODE)).toHaveLength(0);
	});

	it('flags a Type whose only member is in an inactive #If branch', () => {
		const src = 'Public Type T\n#If 0 Then\n    X As Long\n#End If\nEnd Type\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(1);
	});

	it('stays quiet when the member is in an unknown #If branch (cannot prove empty)', () => {
		const src = 'Public Type T\n#If CUSTOM_FLAG Then\n    X As Long\n#End If\nEnd Type\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does not flag an unclosed Type (left to the missing-End-Type diagnostic)', () => {
		expect(byCode(analyzeModule('Public Type T\n'), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate-option (PCEC_006 / CANARY_002)', () => {
	const CODE = 'duplicate-option';

	it('flags a repeated Option Explicit', () => {
		const src = 'Option Explicit\nOption Explicit\n\nPublic Sub T()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('flags two Option Compare even with different arguments', () => {
		expect(byCode(analyzeModule('Option Compare Binary\nOption Compare Text\n'), CODE)).toHaveLength(1);
	});

	it('flags two Option Base', () => {
		expect(byCode(analyzeModule('Option Base 0\nOption Base 1\n'), CODE)).toHaveLength(1);
	});

	it('stays quiet for distinct Option statements', () => {
		const src = 'Option Explicit\nOption Compare Text\nOption Base 1\nOption Private Module\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a single Option Explicit', () => {
		expect(byCode(analyzeModule('Option Explicit\n'), CODE)).toHaveLength(0);
	});

	it('stays quiet for the same Option in mutually-exclusive #If/#Else (unknown constant)', () => {
		const src = '#If CUSTOM_FLAG Then\nOption Explicit\n#Else\nOption Explicit\n#End If\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate-case-else (PCEC_007)', () => {
	const CODE = 'duplicate-case-else';

	it('flags a second Case Else in one Select Case block', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim n As Long\n' +
			'    Select Case n\n' +
			'        Case 1\n' +
			'        Case Else\n' +
			'        Case Else\n' +
			'    End Select\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('Case Else');
	});

	it('stays quiet for a single Case Else', () => {
		const src =
			'Public Sub T()\n    Select Case n\n        Case 1\n        Case Else\n    End Select\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does not collide Case Else across nested Select Case blocks', () => {
		const src =
			'Public Sub T()\n' +
			'    Select Case a\n' +
			'        Case 1\n' +
			'            Select Case b\n' +
			'                Case 2\n' +
			'                Case Else\n' +
			'            End Select\n' +
			'        Case Else\n' +
			'    End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet when one Case Else is in an inactive #If branch', () => {
		const src =
			'Public Sub T()\n' +
			'    Select Case n\n' +
			'        Case 1\n' +
			'#If 0 Then\n' +
			'        Case Else\n' +
			'#End If\n' +
			'        Case Else\n' +
			'    End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for Case Else in mutually-exclusive #If/#Else (unknown constant)', () => {
		const src =
			'Public Sub T()\n' +
			'    Select Case n\n' +
			'        Case 1\n' +
			'#If CUSTOM_FLAG Then\n' +
			'        Case Else\n' +
			'#Else\n' +
			'        Case Else\n' +
			'#End If\n' +
			'    End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - me-outside-object-module (ME_004)', () => {
	const CODE = 'me-outside-object-module';

	it('flags Me used in a standard module', () => {
		const src = 'Public Sub Demo()\n    Debug.Print Me.Name\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('Me');
	});

	it('flags Me as a standalone statement target in a standard module', () => {
		expect(byCode(analyzeModule('Sub T()\n    Me.Activate\nEnd Sub\n'), CODE)).toHaveLength(1);
	});

	it('stays quiet for Me in a class/document/userform module', () => {
		const src = 'Public Sub Demo()\n    Debug.Print Me.Name\nEnd Sub\n';
		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			expect(byCode(analyzeModule(src, { moduleName: 'C', moduleKind }), CODE)).toHaveLength(0);
		}
	});

	it('stays quiet for a member named Me (after a dot) in a standard module', () => {
		expect(byCode(analyzeModule('Sub T()\n    Debug.Print obj.Me\nEnd Sub\n'), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - too-many-parameters (ARG_LIMIT_001B)', () => {
	const CODE = 'too-many-parameters';
	const params = (n: number) => Array.from({ length: n }, (_, i) => `a${i + 1}`).join(', ');

	it('flags a procedure with more than 60 parameters', () => {
		const src = `Public Sub Demo(${params(61)})\nEnd Sub\n`;
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('stays quiet at exactly 60 parameters (the documented maximum)', () => {
		const src = `Public Sub Demo(${params(60)})\nEnd Sub\n`;
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for an ordinary procedure', () => {
		expect(byCode(analyzeModule('Sub T(a As Long, b As Long)\nEnd Sub\n'), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - malformed declarations/statements (parser error-emission)', () => {
	it('flags a leading-underscore identifier (invalid-identifier-start)', () => {
		expect(byCode(analyzeModule('Sub T()\n    Dim _value As Long\nEnd Sub\n'), 'invalid-identifier-start')).toHaveLength(1);
	});

	it('exempts a bracketed name and a mid-word underscore', () => {
		expect(byCode(analyzeModule('Sub T()\n    Dim [_weird] As Long\nEnd Sub\n'), 'invalid-identifier-start')).toHaveLength(0);
		expect(byCode(analyzeModule('Sub T()\n    Dim good_name As Long\nEnd Sub\n'), 'invalid-identifier-start')).toHaveLength(0);
		expect(byCode(analyzeModule('Sub T()\n    Dim good_name As Long\nEnd Sub\n'), 'invalid-identifier-character')).toHaveLength(0);
	});

	it('flags hyphen and dot in a declared name (invalid-identifier-character)', () => {
		expect(byCode(analyzeModule('Sub T()\n    Dim user-name As String\nEnd Sub\n'), 'invalid-identifier-character')).toHaveLength(1);
		expect(byCode(analyzeModule('Sub T()\n    Dim bad.name As Long\nEnd Sub\n'), 'invalid-identifier-character')).toHaveLength(1);
	});

	it('flags assignment to a literal, but not line-numbered statements', () => {
		expect(byCode(analyzeModule('Sub T()\n    1 = x\nEnd Sub\n'), 'invalid-assignment-target')).toHaveLength(1);
		expect(byCode(analyzeModule('Sub T()\n    "s" = x\nEnd Sub\n'), 'invalid-assignment-target')).toHaveLength(1);
		expect(byCode(analyzeModule('Sub T()\n    x = 1\nEnd Sub\n'), 'invalid-assignment-target')).toHaveLength(0);
		expect(byCode(analyzeModule('Sub T()\n    1 x = 5\nEnd Sub\n'), 'invalid-assignment-target')).toHaveLength(0);
	});

	it('flags an Open statement missing For, but not a valid Open', () => {
		expect(byCode(analyzeModule('Sub T()\n    Open "C:\\f.txt" Output #1\nEnd Sub\n'), 'open-missing-for')).toHaveLength(1);
		expect(byCode(analyzeModule('Sub T()\n    Open "C:\\f.txt" For Output As #1\nEnd Sub\n'), 'open-missing-for')).toHaveLength(0);
	});

	it('flags TypeOf with no operand, but not a valid TypeOf or an inactive branch', () => {
		expect(byCode(analyzeModule('Sub T(o As Object)\n    If TypeOf Is Worksheet Then\n    End If\nEnd Sub\n'), 'typeof-missing-operand')).toHaveLength(1);
		expect(byCode(analyzeModule('Sub T(o As Object)\n    If TypeOf o Is Worksheet Then\n    End If\nEnd Sub\n'), 'typeof-missing-operand')).toHaveLength(0);
		expect(byCode(analyzeModule('#If 0 Then\nSub T(o As Object)\n    If TypeOf Is Worksheet Then\n    End If\nEnd Sub\n#End If\n'), 'typeof-missing-operand')).toHaveLength(0);
	});
});

describe('analyzeModule - UDT parameter constraints (SIG_007 / API_VIS_003)', () => {
	const T = 'Private Type Customer\n    Id As Long\nEnd Type\n';
	const OPT = 'optional-udt-parameter';
	const BV = 'byval-udt-parameter';

	it('flags an Optional UDT parameter (ByVal or ByRef)', () => {
		expect(byCode(analyzeModule(T + 'Sub D(Optional ByVal v As Customer)\nEnd Sub\n'), OPT)).toHaveLength(1);
		expect(byCode(analyzeModule(T + 'Sub D(Optional ByRef v As Customer)\nEnd Sub\n'), OPT)).toHaveLength(1);
	});

	it('flags a ByVal UDT parameter (must be ByRef)', () => {
		const hits = byCode(analyzeModule(T + 'Sub D(ByVal v As Customer)\nEnd Sub\n'), BV);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('reports only the Optional diagnostic for an Optional ByVal UDT param (matches VBE)', () => {
		const src = T + 'Sub D(Optional ByVal v As Customer)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), OPT)).toHaveLength(1);
		expect(byCode(analyzeModule(src), BV)).toHaveLength(0);
	});

	it('stays quiet for a ByRef or implicit UDT parameter (those are valid)', () => {
		expect(byCode(analyzeModule(T + 'Sub D(ByRef v As Customer)\nEnd Sub\n'), BV)).toHaveLength(0);
		expect(byCode(analyzeModule(T + 'Sub D(v As Customer)\nEnd Sub\n'), BV)).toHaveLength(0);
	});

	it('stays quiet for scalar parameters and unknown (non-module-UDT) types', () => {
		expect(byCode(analyzeModule(T + 'Sub D(ByVal n As Long)\nEnd Sub\n'), BV)).toHaveLength(0);
		expect(byCode(analyzeModule(T + 'Sub D(Optional ByVal n As Long)\nEnd Sub\n'), OPT)).toHaveLength(0);
		expect(byCode(analyzeModule('Sub D(ByVal v As Widget)\nEnd Sub\n'), BV)).toHaveLength(0);
		expect(byCode(analyzeModule('Sub D(Optional ByVal v As Widget)\nEnd Sub\n'), OPT)).toHaveLength(0);
	});
});

describe('analyzeModule - too-many-array-dimensions (ARRAY_LIMIT_001B)', () => {
	const CODE = 'too-many-array-dimensions';
	const dims = (n: number) => Array.from({ length: n }, () => '1 To 1').join(', ');

	it('flags an array with more than 60 dimensions', () => {
		const src = `Sub T()\n    Dim a(${dims(61)}) As Long\nEnd Sub\n`;
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('stays quiet at exactly 60 dimensions', () => {
		expect(byCode(analyzeModule(`Sub T()\n    Dim a(${dims(60)}) As Long\nEnd Sub\n`), CODE)).toHaveLength(0);
	});

	it('stays quiet for an ordinary array', () => {
		expect(byCode(analyzeModule('Dim a(1 To 3, 0 To 5) As Long\n'), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - identifier-too-long (NAME_LIMIT_001B)', () => {
	const CODE = 'identifier-too-long';

	it('flags a declared name longer than 255 characters', () => {
		const src = `Sub T()\n    Dim ${'a'.repeat(256)} As Long\nEnd Sub\n`;
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('flags an over-long procedure name', () => {
		expect(byCode(analyzeModule(`Sub ${'p'.repeat(256)}()\nEnd Sub\n`), CODE)).toHaveLength(1);
	});

	it('stays quiet at exactly 255 characters', () => {
		expect(byCode(analyzeModule(`Sub T()\n    Dim ${'a'.repeat(255)} As Long\nEnd Sub\n`), CODE)).toHaveLength(0);
	});

	it('stays quiet for ordinary names', () => {
		expect(byCode(analyzeModule('Sub T()\n    Dim count As Long\nEnd Sub\n'), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - else-without-if (CTRL_IF_004)', () => {
	const CODE = 'else-without-if';

	it('flags a stray Else with no enclosing If block', () => {
		const src = 'Public Sub Demo()\n    Else\n        Debug.Print "bad"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('Else');
	});

	it('flags a stray ElseIf', () => {
		expect(byCode(analyzeModule('Sub T()\n    ElseIf y Then\nEnd Sub\n'), CODE)).toHaveLength(1);
	});

	it('flags an Else inside a non-If block (e.g. a For loop)', () => {
		expect(byCode(analyzeModule('Sub T()\n    For i = 1 To 2\n        Else\n    Next\nEnd Sub\n'), CODE)).toHaveLength(1);
	});

	it('stays quiet for a normal block If with ElseIf and Else', () => {
		const src = 'Sub T()\n    If x Then\n        a = 1\n    ElseIf y Then\n        c = 1\n    Else\n        b = 2\n    End If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a single-line If with an inline Else', () => {
		expect(byCode(analyzeModule('Sub T()\n    If x Then a = 1 Else b = 2\nEnd Sub\n'), CODE)).toHaveLength(0);
	});

	it('stays quiet for Else inside a nested If block', () => {
		const src =
			'Sub T()\n    If a Then\n        If b Then\n            x = 1\n        Else\n            y = 1\n        End If\n    Else\n        z = 1\n    End If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});
