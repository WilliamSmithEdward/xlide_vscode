import { describe, expect, it } from 'vitest';
import {
	canonicalKeyword,
	CONTEXTUAL_KEYWORDS,
	isReservedIdentifier,
	LITERAL_IDENTIFIERS,
	MARKER_KEYWORDS,
	OPERATOR_IDENTIFIERS,
	RESERVED_NAMES,
	RESERVED_TYPE_IDENTIFIERS,
	SPECIAL_FORMS,
	STATEMENT_KEYWORDS,
	VBA_KEYWORDS,
} from '../src/analyzer/lexer/keywordTable';

describe('keyword table - canonical casing (MS-VBAL 3.3.5.2)', () => {
	it('maps lowercase keywords to canonical capitalization', () => {
		expect(VBA_KEYWORDS['option']).toBe('Option');
		expect(VBA_KEYWORDS['sub']).toBe('Sub');
		expect(VBA_KEYWORDS['addressof']).toBe('AddressOf');
		expect(VBA_KEYWORDS['redim']).toBe('ReDim');
		expect(VBA_KEYWORDS['lset']).toBe('LSet');
		expect(VBA_KEYWORDS['rset']).toBe('RSet');
		expect(VBA_KEYWORDS['withevents']).toBe('WithEvents');
		expect(VBA_KEYWORDS['raiseevent']).toBe('RaiseEvent');
		expect(VBA_KEYWORDS['paramarray']).toBe('ParamArray');
		expect(VBA_KEYWORDS['gosub']).toBe('GoSub');
		expect(VBA_KEYWORDS['goto']).toBe('GoTo');
	});

	it('includes spec reserved-names missing from the roadmap seed', () => {
		expect(VBA_KEYWORDS['cverr']).toBe('CVErr');
		expect(VBA_KEYWORDS['doevents']).toBe('DoEvents');
		expect(VBA_KEYWORDS['abs']).toBe('Abs');
		expect(VBA_KEYWORDS['fix']).toBe('Fix');
		expect(VBA_KEYWORDS['lenb']).toBe('LenB');
		expect(VBA_KEYWORDS['clnglng']).toBe('CLngLng');
		expect(VBA_KEYWORDS['clngptr']).toBe('CLngPtr');
	});

	it('includes spec special-forms', () => {
		expect(VBA_KEYWORDS['lbound']).toBe('LBound');
		expect(VBA_KEYWORDS['ubound']).toBe('UBound');
		expect(VBA_KEYWORDS['array']).toBe('Array');
		expect(VBA_KEYWORDS['inputb']).toBe('InputB');
	});

	it('capitalizes literal identifiers per VBE convention', () => {
		// Spec grammar writes these lowercase; VBE renders them capitalized.
		expect(VBA_KEYWORDS['true']).toBe('True');
		expect(VBA_KEYWORDS['false']).toBe('False');
		expect(VBA_KEYWORDS['nothing']).toBe('Nothing');
		expect(VBA_KEYWORDS['empty']).toBe('Empty');
		expect(VBA_KEYWORDS['null']).toBe('Null');
	});
});

describe('canonicalKeyword', () => {
	it('is case-insensitive (MS-VBAL 3.3.5.2)', () => {
		expect(canonicalKeyword('OPTION')).toBe('Option');
		expect(canonicalKeyword('oPtIoN')).toBe('Option');
		expect(canonicalKeyword('Dim')).toBe('Dim');
	});

	it('returns undefined for non-keywords', () => {
		expect(canonicalKeyword('myVariable')).toBeUndefined();
		expect(canonicalKeyword('foo')).toBeUndefined();
		expect(canonicalKeyword('')).toBeUndefined();
	});

	it('resolves contextual keywords', () => {
		expect(canonicalKeyword('explicit')).toBe('Explicit');
		expect(canonicalKeyword('property')).toBe('Property');
		expect(canonicalKeyword('lib')).toBe('Lib');
		expect(canonicalKeyword('alias')).toBe('Alias');
		expect(canonicalKeyword('step')).toBe('Step');
	});
});

describe('isReservedIdentifier (MS-VBAL 3.3.5.2)', () => {
	it('recognizes reserved identifiers', () => {
		expect(isReservedIdentifier('Sub')).toBe(true);
		expect(isReservedIdentifier('sub')).toBe(true);
		expect(isReservedIdentifier('Long')).toBe(true);
		expect(isReservedIdentifier('True')).toBe(true);
		expect(isReservedIdentifier('AddressOf')).toBe(true);
	});

	it('rejects contextual keywords (not in the reserved-identifier set)', () => {
		// "Explicit", "Property", "Lib", "Alias" are VBE-convention keywords but
		// are NOT reserved identifiers per the spec grammar.
		expect(isReservedIdentifier('Explicit')).toBe(false);
		expect(isReservedIdentifier('Property')).toBe(false);
		expect(isReservedIdentifier('Lib')).toBe(false);
		expect(isReservedIdentifier('Step')).toBe(false);
	});

	it('rejects ordinary identifiers', () => {
		expect(isReservedIdentifier('myVar')).toBe(false);
		expect(isReservedIdentifier('Worksheet')).toBe(false);
	});
});

describe('keyword table completeness', () => {
	it('contains every reserved-identifier category', () => {
		const categories = [
			STATEMENT_KEYWORDS,
			MARKER_KEYWORDS,
			OPERATOR_IDENTIFIERS,
			RESERVED_NAMES,
			SPECIAL_FORMS,
			RESERVED_TYPE_IDENTIFIERS,
			LITERAL_IDENTIFIERS,
		];
		for (const category of categories) {
			for (const word of category) {
				expect(canonicalKeyword(word)).toBe(word);
			}
		}
	});

	it('contains every contextual keyword', () => {
		for (const word of CONTEXTUAL_KEYWORDS) {
			expect(canonicalKeyword(word)).toBe(word);
		}
	});

	it('every map value is its own canonical form (idempotent)', () => {
		for (const [key, value] of Object.entries(VBA_KEYWORDS)) {
			expect(key).toBe(key.toLowerCase());
			expect(canonicalKeyword(value)).toBe(value);
		}
	});
});
