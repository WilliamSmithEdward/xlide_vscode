import { describe, expect, it, vi } from 'vitest';
import { createConditionalActivityTracker, indexConditionalCompilation } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';

describe('conditional constant replay', () => {
	it('does not copy preceding constants at each directive', () => {
		const count = 400;
		const source = Array.from({ length: count }, (_, i) => '#Const FLAG' + i + ' = ' + (i === 0 ? '1' : 'FLAG' + (i - 1) + ' + 1') + '\n').join('') + '#If FLAG399 = 400 Then\nDebug.Print 1\n#End If';
		const module = parseModule(source);
		let entries = 0;
		const original = Map.prototype[Symbol.iterator];
		const spy = vi.spyOn(Map.prototype, Symbol.iterator).mockImplementation(function* (this: Map<unknown, unknown>) {
			for (const entry of original.call(this)) { entries++; yield entry; }
		});
		try {
			expect(indexConditionalCompilation(module).constants[count - 1].value).toBe(count);
			const tracker = createConditionalActivityTracker(module)!;
			const start = source.indexOf('Debug.Print');
			expect(tracker.activityForSpan({ start, end: start + 1 })).toBe('active');
			expect(entries).toBeLessThan(count * 10);
		} finally { spy.mockRestore(); }
	});

	it('preserves the distinct missing-name policies of index and activity replay', () => {
		const source = '#Const NEXT = MISSING + 1\n#If NEXT = 1 Then\nDebug.Print 1\n#End If';
		const module = parseModule(source);
		expect(indexConditionalCompilation(module).constants[0].value).toBe(1);
		const start = source.indexOf('Debug.Print');
		expect(createConditionalActivityTracker(module)!.activityForSpan({ start, end: start + 1 })).toBe('unknown');
		expect(createConditionalActivityTracker(module, { projectConstants: {} })!.activityForSpan({ start, end: start + 1 })).toBe('active');
	});

	it.each(['0', 'False', '""', 'Null'])('lets module constants shadow compiler/project values with %s', value => {
		const source = '#Const SHARED = ' + value + '\n#Const COPIED = SHARED\n';
		const env = { compilerConstants: { shared: 99 }, projectConstants: { Shared: 42 } };
		const result = indexConditionalCompilation(parseModule(source), env);
		expect(result.constants[1].value).toEqual(result.constants[0].value);
		expect(result.constants[1].value).not.toBe(99);
		expect(result.constants[1].value).not.toBe(42);
		expect(env.projectConstants.Shared).toBe(42);
	});
});
