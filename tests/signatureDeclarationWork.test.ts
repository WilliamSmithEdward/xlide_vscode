import { beforeEach, expect, it, vi } from 'vitest';
import { resolveSignatureHelp } from '../src/analyzer/signature/signatureHelp';
const work = vi.hoisted(() => ({ reads: 0 }));
vi.mock('../src/analyzer/parser/parseModule', async () => {
	const actual = await vi.importActual<typeof import('../src/analyzer/parser/parseModule')>('../src/analyzer/parser/parseModule');
	const views = new WeakMap<object, ReturnType<typeof actual.parseModule>>();
	return { ...actual, parseModule(source: string) {
		const module = actual.parseModule(source);
		let view = views.get(module);
		if (!view) {
			const members = new Proxy(module.members, { get(target, key, receiver) {
				if (key === Symbol.iterator) return function* () { for (const member of target) { work.reads++; yield member; } };
				return Reflect.get(target, key, receiver);
			} });
			view = { ...module, members };
			views.set(module, view);
		}
		return view;
	} };
});
beforeEach(() => { work.reads = 0; });
const caller = 'Sub Caller()\ntArGeT(\nEnd Sub';
const offset = caller.indexOf('tArGeT(') + 'tArGeT('.length;
for (const count of [10, 100, 1000]) it(`indexes ${count} local procedures once across repeated signature requests`, () => {
	const moduleSource = Array.from({ length: count }, (_, i) => `Sub P${i}()\nEnd Sub\n`).join('') + 'Sub Target(ByVal value As Long)\nEnd Sub\n';
	for (let i = 0; i < 10; i++) expect(resolveSignatureHelp(caller, offset, { moduleSource })).toEqual({ label: 'Target(value As Long)', parameters: [{ label: 'value As Long', documentation: undefined }], activeParameter: 0, documentation: undefined, details: undefined });
	expect(work.reads).toBeLessThanOrEqual(2 * (count + 1));
});
it('preserves first declaration wins and updates changed module sources', () => {
	for (const type of ['Long', 'String', 'Long']) {
		const moduleSource = `Sub TARGET(ByVal value As ${type})\nEnd Sub\nSub target(ByVal ignored As Object)\nEnd Sub\n`;
		expect(resolveSignatureHelp(caller, offset, { moduleSource })?.label).toBe(`TARGET(value As ${type})`);
	}
});
