import { describe, it, expect } from 'vitest';
import {
	resolveRuntimeFunction,
	VBA_RUNTIME_FUNCTIONS,
	resolveIdentifierCompletions,
} from '../src/analyzer';

describe('VBA runtime metadata', () => {
	it('resolves built-ins case-insensitively', () => {
		expect(resolveRuntimeFunction('MsgBox')?.name).toBe('MsgBox');
		expect(resolveRuntimeFunction('msgbox')?.name).toBe('MsgBox');
		expect(resolveRuntimeFunction('LEFT')?.returns).toBe('String');
		expect(resolveRuntimeFunction('CLng')?.returns).toBe('Long');
		expect(resolveRuntimeFunction('NotARealFunction')).toBeUndefined();
	});

	it('every entry has a signature and is marked verified', () => {
		for (const f of VBA_RUNTIME_FUNCTIONS) {
			expect(f.name.length).toBeGreaterThan(0);
			expect(f.signature.length).toBeGreaterThan(0);
			expect(f.source).toBe('verified');
			if (f.kind === 'function') {
				expect(f.signature).toContain('As ');
			}
		}
	});

	it('excludes intrinsic data-type names to avoid type/function ambiguity', () => {
		for (const name of ['String', 'Date', 'Time', 'Error']) {
			expect(resolveRuntimeFunction(name)).toBeUndefined();
		}
	});

	it('includes verified gap-filled built-ins from the VBA library', () => {
		expect(resolveRuntimeFunction('StrReverse')?.returns).toBe('String');
		expect(resolveRuntimeFunction('FormatCurrency')?.kind).toBe('function');
		expect(resolveRuntimeFunction('Dir')?.returns).toBe('String');
		expect(resolveRuntimeFunction('FreeFile')?.returns).toBe('Integer');
		expect(resolveRuntimeFunction('Pmt')?.returns).toBe('Double');
		expect(resolveRuntimeFunction('CallByName')?.name).toBe('CallByName');
		// File-system commands are statements, not value-returning functions.
		expect(resolveRuntimeFunction('Kill')?.kind).toBe('statement');
		expect(resolveRuntimeFunction('MkDir')?.kind).toBe('statement');
	});
});

describe('identifier completion - runtime built-ins', () => {
	const src = 'Sub T()\n    \nEnd Sub\n';
	const offset = src.indexOf('    \n') + 4;

	it('offers built-in functions at a statement start', () => {
		const names = resolveIdentifierCompletions(src, offset, {
			moduleName: 'M',
		}).map((c) => c.name);
		expect(names).toContain('MsgBox');
		expect(names).toContain('Array');
		expect(names).toContain('RGB');
	});

	it('filters built-ins by the typed prefix', () => {
		const typed = 'Sub T()\n    Msg\nEnd Sub\n';
		const off = typed.indexOf('Msg') + 3;
		const names = resolveIdentifierCompletions(typed, off, {
			moduleName: 'M',
		}).map((c) => c.name);
		expect(names).toContain('MsgBox');
		expect(names).not.toContain('Array');
	});

	it('can be disabled via includeRuntime', () => {
		const names = resolveIdentifierCompletions(src, offset, {
			moduleName: 'M',
			includeRuntime: false,
		}).map((c) => c.name);
		expect(names).not.toContain('MsgBox');
	});
});
