import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const hosts = ['excel', 'word', 'powerpoint', 'access'] as const;

describe('statement forms with an absent Call target (issue #853)', () => {
	for (const host of hosts) {
		it.each(['Call:', 'Call', '100 Call', "Call ' still typing", 'If True Then Call', '100', 'L:', '100 L:', 'Call Foo', 'L: Call Foo'])('keeps checking after %s in ' + host, (line) => {
			const source = 'Private Sub Foo()\nEnd Sub\nSub Main()\nDim x As Variant\n' + line + '\nx = Foo\nEnd Sub\n';
			const errors: string[] = [];
			const diagnostics = analyzeModule(source, { host, onInternalError: (error) => errors.push(String(error)) });
			expect(errors).toEqual([]);
			expectDiagnostic(source, diagnostics, 'sub-used-as-value', { span: 'Foo', message: 'returns nothing' });
			if (line === 'Call:') {
				expectDiagnostic(source, diagnostics, 'malformed-statement', { span: 'Call', message: 'reserved word' });
			}
			if (line === 'L:' || line === '100' || line === '100 L:') {
				expect(byCode(diagnostics, 'malformed-statement')).toEqual([]);
			}
		});

		it('handles Call at the end of incomplete editor input in ' + host, () => {
			const errors: string[] = [];
			analyzeModule('Sub Main()\nCall', { host, onInternalError: (error) => errors.push(String(error)) });
			expect(errors).toEqual([]);
		});
	}
});
