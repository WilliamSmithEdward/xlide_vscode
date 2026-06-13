// Diagnostics tests: parameter-default-not-constant rule.
//
// Optional parameter defaults must be constant expressions (MS-VBAL 5.3.1.5;
// VBE "Constant expression required"). Only provably non-constant defaults —
// calls, New, AddressOf — are flagged; anything that may be a constant stays quiet.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

const CODE = 'parameter-default-not-constant';

function hits(src: string) {
	return byCode(analyzeModule(src), CODE);
}

describe('analyzeModule - parameter-default-not-constant', () => {
	it('flags a function-call default', () => {
		const src = 'Sub T(Optional ByVal n As Long = GetDefault())\nEnd Sub\n';
		const found = hits(src);
		expect(found).toHaveLength(1);
		expect(found[0].severity).toBe('error');
		expect(spanText(src, found[0])).toBe('GetDefault()');
	});

	it('flags a New default and an array-index default', () => {
		expect(hits('Sub T(Optional x As Variant = New Collection)\nEnd Sub\n')).toHaveLength(1);
		expect(hits('Sub T(Optional ByVal n As Long = arr(0))\nEnd Sub\n')).toHaveLength(1);
	});

	it('stays quiet for constant defaults (literal, arithmetic, grouped, string)', () => {
		expect(hits('Sub T(Optional ByVal n As Long = 5)\nEnd Sub\n')).toHaveLength(0);
		expect(hits('Sub T(Optional ByVal n As Long = 1 + 2)\nEnd Sub\n')).toHaveLength(0);
		expect(hits('Sub T(Optional ByVal n As Long = (1 + 2))\nEnd Sub\n')).toHaveLength(0);
		expect(hits('Sub T(Optional ByVal s As String = "x")\nEnd Sub\n')).toHaveLength(0);
	});

	it('stays quiet for a bare or qualified identifier (may be a Const/Enum member)', () => {
		expect(hits('Sub T(Optional ByVal n As Long = MAX)\nEnd Sub\n')).toHaveLength(0);
		expect(hits('Sub T(Optional ByVal n As Long = Module1.MAX)\nEnd Sub\n')).toHaveLength(0);
		expect(hits('Sub T(Optional ByVal n As Long = vbCrLf)\nEnd Sub\n')).toHaveLength(0);
	});

	it('does not double-report on an object parameter (owned by the type rule)', () => {
		// Object-typed param defaults are handled by parameter-default-type-mismatch.
		expect(hits('Sub T(Optional o As Worksheet = Nothing)\nEnd Sub\n')).toHaveLength(0);
	});

	it('stays quiet for a parameter with no default', () => {
		expect(hits('Sub T(ByVal n As Long)\nEnd Sub\n')).toHaveLength(0);
	});
});
