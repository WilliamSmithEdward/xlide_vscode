// github.com/WilliamSmithEdward/xlide_vscode/issues/61.
//
// Extract Variable needs the type of a selected expression and whether
// assigning it takes `Set`. Both live in the binder; nothing exposed them for a
// SPAN. `complete` matters as much as the type: a selection of `1 +` is not an
// expression, and the refactoring must decline rather than emit a line that
// does not parse.

import { describe, expect, it } from 'vitest';
import { resolveExpressionType } from '../src/analyzer';

const HEADER = [
	'Option Explicit',
	'Public Type TPoint',
	'    X As Long',
	'End Type',
	'Sub T()',
	'    Dim num As Long',
	'    Dim txt As String',
	'    Dim col As Collection',
	'    Dim var As Variant',
	'    Dim obj As Object',
	'    Dim pt As TPoint',
];

/** Resolves `needle` inside a procedure body holding `statement`. */
function typeOf(statement: string, needle: string) {
	const source = [...HEADER, `    ${statement}`, 'End Sub'].join('\r\n') + '\r\n';
	const start = source.lastIndexOf(needle);
	expect(start, needle).toBeGreaterThanOrEqual(0);
	return resolveExpressionType(source, { start, end: start + needle.length }, { moduleName: 'M' });
}

describe('the declared type of an expression', () => {
	it('answers for a local of each shape', () => {
		expect(typeOf('num = 1', 'num')).toMatchObject({ type: 'Long', isObject: false });
		expect(typeOf('txt = "a"', 'txt')).toMatchObject({ type: 'String', isObject: false });
		expect(typeOf('var = 1', 'var')).toMatchObject({ type: 'Variant', isObject: false });
	});

	it('answers for literals and arithmetic', () => {
		expect(typeOf('txt = "hello"', '"hello"')).toMatchObject({ type: 'String', isObject: false });
		expect(typeOf('num = 1 + 2', '1 + 2')).toMatchObject({ isObject: false, complete: true });
	});
});

// github.com/WilliamSmithEdward/xlide_vscode/issues/64. VBA types a
// whole-number literal as Integer or Long, never Double. The shared inference
// widens every numeric literal to Double, which is right for CHECKING
// compatibility and wrong for a caller about to write the type into a `Dim` -
// and every caller of this function is doing exactly that. Asked of the parse,
// so the suffixed, hex and parenthesised forms all agree.
describe('a whole-number literal is a Long', () => {
	it('answers Long for every spelling of one', () => {
		for (const literal of ['10', '10&', '&H10', '(10)', '-10']) {
			expect(typeOf(`num = ${literal}`, literal), literal).toMatchObject({ type: 'Long' });
		}
	});

	it('leaves anything else to the shared inference', () => {
		expect(typeOf('num = 1.5', '1.5')).toMatchObject({ type: 'Double' });
		expect(typeOf('num = 1 + 2', '1 + 2')).toMatchObject({ type: 'Double' });
	});
});

describe('whether the assignment needs Set', () => {
	it('says yes for an object', () => {
		expect(typeOf('Set obj = Nothing', 'obj')).toMatchObject({ type: 'Object', isObject: true });
		expect(typeOf('Set col = Nothing', 'col')).toMatchObject({ type: 'Collection', isObject: true });
	});

	it('says yes for `New`, whatever the class is', () => {
		expect(typeOf('Set col = New Collection', 'New Collection'))
			.toMatchObject({ type: 'Collection', isObject: true });
	});

	// A user-defined Type is assigned with a plain `=`, so "not a scalar" is the
	// wrong test and would emit a `Set` that does not compile.
	it('says no for a user-defined Type', () => {
		expect(typeOf('pt.X = 1', 'pt')).toMatchObject({ type: 'TPoint', isObject: false });
	});

	it('says no for Variant, which takes either form', () => {
		expect(typeOf('var = 1', 'var')).toMatchObject({ isObject: false });
	});
});

describe('whether the span holds a whole expression', () => {
	it('accepts a complete one', () => {
		expect(typeOf('num = (1 + 2)', '(1 + 2)')?.complete).toBe(true);
		expect(typeOf('num = 1 + 2', '1 + 2')?.complete).toBe(true);
	});

	it('rejects a trailing operator and an unbalanced paren', () => {
		expect(typeOf('num = 1 + 2', '1 +')?.complete).toBe(false);
		expect(typeOf('num = (1 + 2)', '(1 + 2')?.complete).toBe(false);
	});

	it('returns nothing at all for a span holding no expression', () => {
		const source = [...HEADER, '    num = 1', 'End Sub'].join('\r\n') + '\r\n';
		expect(resolveExpressionType(source, { start: 0, end: 0 }, { moduleName: 'M' })).toBeUndefined();
	});
});
