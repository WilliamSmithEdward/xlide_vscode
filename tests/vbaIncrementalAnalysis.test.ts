import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// Incremental rule re-analysis: body-only edits re-walk just the dirty
// procedure and must produce diagnostics identical to a full pass; any
// envelope (declarations/signatures/directives) or fingerprint change must
// fall back to a full pass.

const BASE = [
	'Option Explicit',
	'Private mCount As Long',
	'',
	'Sub Alpha()',
	'    Dim a As Long',
	'    a = 1',
	'End Sub',
	'',
	'Sub Beta()',
	'    undeclaredBeta = 2',
	'End Sub',
	'',
	'Function Gamma() As Long',
	'    Gamma = mCount',
	'End Function',
].join('\n');

const FP = ['fp-a'] as const;

function run(source: string, state?: ReturnType<typeof analyzeVbaModuleSource>['rulesIncrementalState'], fingerprint: readonly unknown[] = FP) {
	return analyzeVbaModuleSource({
		source,
		moduleName: 'Module1',
		knownIdentifiers: new Set<string>(),
		rulesIncremental: { state, fingerprint },
	});
}

function key(r: ReturnType<typeof analyzeVbaModuleSource>): string {
	return r.diagnostics
		.map((d) => `${d.code}:${d.span.start}:${d.span.end}:${d.severity}:${d.message}`)
		.sort()
		.join('\n');
}

function full(source: string): ReturnType<typeof analyzeVbaModuleSource> {
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', knownIdentifiers: new Set<string>() });
}

describe('incremental rule re-analysis', () => {
	it('re-analyzes only a changed body and matches the full pass exactly', () => {
		const base = run(BASE);
		expect(base.rulesIncrementalMode).toBe('full');
		// Introduce a new error inside Alpha's body.
		const edited = BASE.replace('    a = 1', '    a = 1\n    undeclaredAlpha = 3');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('incremental');
		expect(key(incr)).toBe(key(full(edited)));
		// The pre-existing error in the untouched Beta must survive the splice.
		expect(incr.diagnostics.some((d) => d.message.includes('undeclaredBeta'))).toBe(true);
		expect(incr.diagnostics.some((d) => d.message.includes('undeclaredAlpha'))).toBe(true);
	});

	it('clears a fixed error in the edited body and keeps others', () => {
		const base = run(BASE);
		const edited = BASE.replace('    undeclaredBeta = 2', '    mCount = 2');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('incremental');
		expect(key(incr)).toBe(key(full(edited)));
		expect(incr.diagnostics.some((d) => d.message.includes('undeclaredBeta'))).toBe(false);
	});

	it('chains state across successive edits', () => {
		let source = BASE;
		let r = run(source);
		for (let i = 0; i < 3; i += 1) {
			source = source.replace('    a = 1', `    a = 1\n    a = a + ${i}`);
			r = run(source, r.rulesIncrementalState);
			expect(r.rulesIncrementalMode).toBe('incremental');
			expect(key(r)).toBe(key(full(source)));
		}
	});

	it('falls back to full on a signature change', () => {
		const base = run(BASE);
		const edited = BASE.replace('Sub Alpha()', 'Sub Alpha(ByVal n As Long)');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('full');
		expect(key(incr)).toBe(key(full(edited)));
	});

	it('falls back to full on a declaration-section change', () => {
		const base = run(BASE);
		const edited = BASE.replace('Private mCount As Long', 'Private mCount As Long\nPrivate mOther As String');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('full');
		expect(key(incr)).toBe(key(full(edited)));
	});

	it('falls back to full when a procedure is added', () => {
		const base = run(BASE);
		const edited = `${BASE}\n\nSub Delta()\nEnd Sub`;
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('full');
		expect(key(incr)).toBe(key(full(edited)));
	});

	it('falls back to full on a fingerprint change (cross-module inputs moved)', () => {
		const base = run(BASE);
		const edited = BASE.replace('    a = 1', '    a = 2');
		const incr = run(edited, base.rulesIncrementalState, ['fp-b']);
		expect(incr.rulesIncrementalMode).toBe('full');
		expect(key(incr)).toBe(key(full(edited)));
	});
});
