import { beforeEach, expect, it, vi } from 'vitest';
import { implementInterface } from '../src/analyzer/refactor/implementInterface';
const work = vi.hoisted(() => ({ sources: [] as string[] }));
vi.mock('../src/analyzer/parser/parseModule', async () => {
	const actual = await vi.importActual<typeof import('../src/analyzer/parser/parseModule')>('../src/analyzer/parser/parseModule');
	return { ...actual, parseModule(...args: Parameters<typeof actual.parseModule>) {
		work.sources.push(args[0]);
		return actual.parseModule(...args);
	} };
});
beforeEach(() => { work.sources.length = 0; });
for (const count of [10, 1000]) for (const eol of ['\n', '\r\n', '\r']) {
	it(`refuses a class without Implements before parsing ${count} procedures with ${JSON.stringify(eol)}`, () => {
		const source = Array.from({ length: count }, (_, i) => `Sub P${i}()${eol}End Sub${eol}`).join('');
		expect(implementInterface({ source, moduleSources: {} })).toEqual({ ok: false, reason: 'This class implements no interface. Add an `Implements` statement first.' });
		expect(work.sources).toEqual([]);
	});
}
it('defers class parsing across choice, membership, missing-source and empty-interface refusals', () => {
	const cases = [
		{ source: 'Implements IFirst\nImplements ISecond', moduleSources: {}, reason: 'This class implements IFirst, ISecond. Say which one to implement.' },
		{ source: 'Implements IFirst', interfaceName: 'Other', moduleSources: {}, reason: "This class does not implement 'Other'." },
		{ source: 'Implements IFirst', moduleSources: {}, reason: "The project has no module called 'IFirst'." },
		{ source: 'Implements IFirst', moduleSources: { IFirst: 'Private Sub Hidden()\nEnd Sub' }, reason: "'IFirst' has no public members to implement." },
	];
	for (const { reason, ...input } of cases) {
		work.sources.length = 0;
		expect(implementInterface(input)).toEqual({ ok: false, reason });
		expect(work.sources).toEqual(input.moduleSources.IFirst ? [input.moduleSources.IFirst] : []);
	}
});
it('still parses both modules for generation and already-implemented checks', () => {
	const interfaceSource = 'Public Sub Work(ByVal Input As Long)\nEnd Sub';
	const source = 'Implements IJob\n';
	const expected = { ok: true, title: "Implement 1 member of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: "\nPrivate Sub IJob_Work(ByVal Input As Long)\n    Err.Raise 5 'TODO: implement this interface member\nEnd Sub\n" }] };
	expect(implementInterface({ source, moduleSources: { IJob: interfaceSource } })).toEqual(expected);
	expect(new Set(work.sources)).toEqual(new Set([source, interfaceSource]));
	work.sources.length = 0;
	const complete = source + 'Private Sub IJob_Work(ByVal Input As Long)\nEnd Sub';
	expect(implementInterface({ source: complete, moduleSources: { IJob: interfaceSource } })).toEqual({ ok: false, reason: "'IJob' is already implemented in full." });
	expect(new Set(work.sources)).toEqual(new Set([complete, interfaceSource]));
});
