// Diagnostics tests: const-value-not-constant and enum-member-not-constant.
//
// Const declaration values (MS-VBAL 5.2.4) and Enum member initializers
// (MS-VBAL 5.2.3.4) must be constant expressions; VBE rejects a non-constant
// with "Constant expression required". Only provably non-constant elements —
// a call (name(...)), New, or AddressOf — are flagged. Anything that may be a
// constant (bare/qualified identifiers, literals, string concat, arithmetic,
// grouping) stays quiet, keeping both rules no-false-positive.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

const CONST_CODE = 'const-value-not-constant';
const ENUM_CODE = 'enum-member-not-constant';

function constHits(src: string) {
	return byCode(analyzeModule(src), CONST_CODE);
}

function enumHits(src: string) {
	return byCode(analyzeModule(src), ENUM_CODE);
}

describe('analyzeModule - const-value-not-constant', () => {
	it('flags a module-level Const whose value is a function/array call', () => {
		const src = 'Const X = GetDefault()\n';
		const found = constHits(src);
		expect(found).toHaveLength(1);
		expect(found[0].severity).toBe('error');
		expect(spanText(src, found[0])).toBe('GetDefault()');
	});

	it('flags New and AddressOf Const values', () => {
		expect(constHits('Const X = New Collection\n')).toHaveLength(1);
		expect(constHits('Const X = AddressOf Foo\n')).toHaveLength(1);
	});

	it('flags a procedure-local Const, including inside a nested block', () => {
		expect(constHits('Sub S()\nConst X = Foo()\nEnd Sub\n')).toHaveLength(1);
		const nested = 'Sub S()\nIf True Then\nConst X = Foo()\nEnd If\nEnd Sub\n';
		expect(constHits(nested)).toHaveLength(1);
	});

	it('flags only the non-constant name in a multi-name Const', () => {
		const src = 'Const A = 1, B = Foo()\n';
		const found = constHits(src);
		expect(found).toHaveLength(1);
		expect(spanText(src, found[0])).toBe('Foo()');
	});

	it('flags a qualified call but not a qualified constant reference', () => {
		expect(constHits('Const X = Module1.Build()\n')).toHaveLength(1);
		expect(constHits('Const X = Module1.MAX\n')).toHaveLength(0);
	});

	it('stays quiet for constant Const values (literal, arithmetic, grouped, hex, string concat)', () => {
		expect(constHits('Const N As Long = 5\n')).toHaveLength(0);
		expect(constHits('Const N As Long = 1 + 2 * 3\n')).toHaveLength(0);
		expect(constHits('Const N As Long = (1 + 2)\n')).toHaveLength(0);
		expect(constHits('Const H As Long = &HFF\n')).toHaveLength(0);
		expect(constHits('Const S As String = "a" & "b"\n')).toHaveLength(0);
	});

	it('stays quiet for a bare identifier value (may be another Const/Enum member)', () => {
		expect(constHits('Const X = OTHER_CONST\n')).toHaveLength(0);
		expect(constHits('Const X = vbCrLf\n')).toHaveLength(0);
	});

	it('stays quiet for operator keywords before a grouping paren (not calls)', () => {
		// And/Or/Not/Mod/Xor/Eqv/Imp/Like lex as keywords but are operators, not
		// callable names: `6 And (3)` is a legal constant expression, not a call.
		expect(constHits('Const X As Long = 6 And (3)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = 4 Or (1)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = Not (0)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = 7 Mod (3)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = 5 Xor (1)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = 5 Eqv (5)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = 1 Imp (0)\n')).toHaveLength(0);
		expect(constHits('Const X As Long = &HFF And (1)\n')).toHaveLength(0);
	});
});

describe('analyzeModule - enum-member-not-constant', () => {
	it('flags an Enum member whose value is a call', () => {
		const src = 'Public Enum E\nA = Build()\nB = 2\nEnd Enum\n';
		const found = enumHits(src);
		expect(found).toHaveLength(1);
		expect(found[0].severity).toBe('error');
		expect(spanText(src, found[0])).toBe('Build()');
	});

	it('flags New and AddressOf Enum member values', () => {
		expect(enumHits('Public Enum E\nA = New Thing\nEnd Enum\n')).toHaveLength(1);
		expect(enumHits('Public Enum E\nA = AddressOf Foo\nEnd Enum\n')).toHaveLength(1);
	});

	it('stays quiet for implicit (auto-numbered) members', () => {
		expect(enumHits('Public Enum E\nA\nB\nC\nEnd Enum\n')).toHaveLength(0);
	});

	it('stays quiet for constant member values (literal, hex, arithmetic on a prior member, bare/qualified ident)', () => {
		expect(enumHits('Public Enum E\nA = 1\nB = 2\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum E\nA = &H10\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum E\nA = 1\nB = A + 1\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum E\nA = OTHER_CONST\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum E\nA = Module1.MAX\nEnd Enum\n')).toHaveLength(0);
	});

	it('flags only the non-constant member when others are constant or implicit', () => {
		const src = 'Public Enum E\nA\nB = 10\nC = Build()\nD = 20\nEnd Enum\n';
		const found = enumHits(src);
		expect(found).toHaveLength(1);
		expect(spanText(src, found[0])).toBe('Build()');
	});

	it('stays quiet for operator keywords before a grouping paren (bitflag patterns)', () => {
		expect(enumHits('Public Enum F\nA = 1\nB = 2\nC = A Or (B)\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum F\nA = 1\nB = 2\nC = A And (B)\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum F\nA = Not (1 + 1)\nEnd Enum\n')).toHaveLength(0);
		expect(enumHits('Public Enum F\nA = 10\nB = A Mod (3)\nEnd Enum\n')).toHaveLength(0);
	});

	it('stays quiet for an Enum body contaminated by inner #If directives (dead branch may hold the value)', () => {
		// The parser does not yet model #If inside an Enum body; a dead-branch
		// member must not be flagged. The block is skipped conservatively.
		const src = 'Public Enum E\nA = 1\n#If 0 Then\nB = Foo()\n#End If\nC = 3\nEnd Enum\n';
		expect(enumHits(src)).toHaveLength(0);
	});
});
