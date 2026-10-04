import { beforeEach, expect, it, vi } from 'vitest';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
const work = vi.hoisted(() => ({ reads: 0 }));
vi.mock('../src/analyzer/symbols/editorModuleSymbols', async () => {
	const actual = await vi.importActual<typeof import('../src/analyzer/symbols/editorModuleSymbols')>('../src/analyzer/symbols/editorModuleSymbols');
	const views = new WeakMap<object, ReturnType<typeof actual.editorModuleSymbols>>();
	const counted = <T>(items: T[]) => new Proxy(items, { get(target, key, receiver) {
		if (typeof key === 'string' && /^\d+$/.test(key)) { work.reads++; }
		return Reflect.get(target, key, receiver);
	} });
	return { ...actual, editorModuleSymbols(moduleName: string, moduleKind: Parameters<typeof actual.editorModuleSymbols>[1], source: string) {
		const original = actual.editorModuleSymbols(moduleName, moduleKind, source);
		let view = views.get(original);
		if (!view) {
			const clones = new Map(original.all.map(symbol => [symbol, { ...symbol, ...(symbol.children ? { children: counted(symbol.children) } : {}) }]));
			view = { ...original, all: counted(original.all.map(symbol => clones.get(symbol)!)), root: { ...original.root, children: counted((original.root.children ?? []).map(symbol => clones.get(symbol) ?? symbol)) } };
			if (source.endsWith("' overlap") || source.endsWith("' touching")) {
				const procedures = view.all.filter(symbol => symbol.kind === 'sub');
				const at = source.lastIndexOf('hovervalue');
				procedures[0].fullSpan = { start: 0, end: source.endsWith("' overlap") ? source.length : at };
				if (source.endsWith("' touching")) { procedures[1].fullSpan = { start: at, end: source.length }; }
			}
			views.set(original, view);
		}
		return view;
	} };
});
beforeEach(() => { work.reads = 0; });
for (const count of [10, 100, 1000]) {
	it(`indexes ${count} preceding procedures for repeated hover requests`, () => {
		const source = Array.from({ length: count }, (_, i) => `Sub P${i}()\nEnd Sub\n`).join('') + 'Sub Target()\nDim HoverValue As Long\nDebug.Print hovervalue\nEnd Sub';
		const start = source.lastIndexOf('hovervalue');
		for (let request = 0; request < 10; request++) {
			expect(resolveHover(source, start + 2)).toEqual({ signature: 'HoverValue As Long', details: ['Local in Target'], span: { start, end: start + 10 } });
		}
		expect(work.reads).toBeLessThanOrEqual(4 * (count + 2));
	});
	it(`indexes ${count} preceding locals for repeated hover requests`, () => {
		const source = 'Sub LocalTarget()\n' + Array.from({ length: count }, (_, i) => `Dim Local${i} As Long\n`).join('') + 'Dim HoverValue As String\nDebug.Print hovervalue\nEnd Sub';
		const start = source.lastIndexOf('hovervalue');
		for (let request = 0; request < 10; request++) {
			expect(resolveHover(source, start + 2)).toEqual({ signature: 'HoverValue As String', details: ['Local in LocalTarget'], span: { start, end: start + 10 } });
		}
		expect(work.reads).toBeLessThanOrEqual(4 * (count + 2));
	});
}

for (const mode of ['overlap', 'touching']) it(`preserves first enclosing scope for ${mode} procedure spans`, () => {
	const source = 'Sub First()\nDim HoverValue As Long\nEnd Sub\nSub Second()\nDim HoverValue As String\nDebug.Print hovervalue\nEnd Sub\n' + "' " + mode;
	const start = source.lastIndexOf('hovervalue');
	expect(resolveHover(source, start)).toEqual({ signature: 'HoverValue As Long', details: ['Local in First'], span: { start, end: start + 10 } });
});
for (const eol of ['\n', '\r\n', '\r']) it(`preserves local/module/enum precedence and first declarations with ${JSON.stringify(eol)}`, () => {
	const source = ['Public HoverValue As String', 'Public hovervalue As Double', 'Enum FirstEnum', 'EnumOnly = 1', 'End Enum', 'Enum SecondEnum', 'enumonly = 2', 'End Enum', 'Sub First()', 'Dim HoverValue As Long', 'Dim hovervalue As Double', 'Debug.Print hovervalue', 'End Sub', 'Sub Second()', 'Debug.Print hovervalue', 'Debug.Print enumonly', 'End Sub'].join(eol);
	const local = source.indexOf('Debug.Print hovervalue') + 'Debug.Print '.length, top = source.lastIndexOf('hovervalue'), member = source.lastIndexOf('enumonly');
	for (let round = 0; round < 3; round++) {
		expect(resolveHover(source, local)).toEqual({ signature: 'HoverValue As Long', details: ['Local in First'], span: { start: local, end: local + 10 } });
		expect(resolveHover(source, top)).toEqual({ signature: 'HoverValue As String', details: ['Declared in Module: Module', 'Visibility: Public'], span: { start: top, end: top + 10 } });
		expect(resolveHover(source, member)).toEqual({ signature: 'EnumOnly', details: ['Member of Enum FirstEnum'], span: { start: member, end: member + 8 } });
	}
});
