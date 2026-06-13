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
