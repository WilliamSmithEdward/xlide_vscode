import { expect, it, vi } from 'vitest';
import { implementInterface } from '../src/analyzer/refactor/implementInterface';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';
const names = ['Δοκιμή', '日本語', 'ไทย', 'कर्म', '[If]', '[Unit Price]'];
const qualified = (name: string) => name.startsWith('[') ? `[IJob_${name.slice(1, -1)}]` : 'IJob_' + name;
for (const name of names) for (const kind of ['Sub', 'Function', 'Property Get']) for (const eol of ['\n', '\r\n', '\r']) {
	it(`renames the declared ${kind} ${name} with ${JSON.stringify(eol)}`, () => {
		const signature = (written: string) => kind === 'Sub' ? `Sub ${written}(ByRef Arg As Long)` : `${kind} ${written}() As Long`;
		const closer = kind === 'Property Get' ? 'Property' : kind;
		const interfaceSource = ['Public ' + signature(name), "    Debug.Print \"private body\"", 'End ' + closer].join(eol);
		const source = 'Implements IJob\n';
		const text = '\nPrivate ' + signature(qualified(name)) + "\n    Err.Raise 5 'TODO: implement this interface member\nEnd " + closer + '\n';
		const result = implementInterface({ source, moduleSources: { IJob: interfaceSource } });
		expect(result).toEqual({ ok: true, title: "Implement 1 member of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: text }] });
		if (!result.ok) { throw new Error(result.reason); }
		const members = parseModule(applyVbaTextEdits(source, result.edits)).members.filter(member => member.kind === 'Procedure');
		expect(members.map(member => member.name)).toEqual(['IJob_' + name.replace(/^\[|\]$/g, '')]);
	});
}
for (const name of names) for (const type of ['Long', 'Object']) {
	it(`renames both property legs for field ${name} As ${type}`, () => {
		const source = 'Implements IJob\n';
		const setter = type === 'Object' ? 'Set' : 'Let';
		const getter = `Private Property Get ${qualified(name)}() As ${type}\n    Err.Raise 5 'TODO: implement this interface member\nEnd Property`;
		const writer = `Private Property ${setter} ${qualified(name)}(ByVal RHS As ${type})\n    Err.Raise 5 'TODO: implement this interface member\nEnd Property`;
		expect(implementInterface({ source, moduleSources: { IJob: `Public ${name} As ${type}` } })).toEqual({ ok: true, title: "Implement 2 members of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: '\n' + getter + '\n\n' + writer + '\n' }] });
	});
}
it('changes only the declaration name, preserving matching parameters and literal/comment text', () => {
	const source = 'Implements IJob\n';
	const interfaceSource = 'Public Sub [If](ByVal [If] As Long, Optional ByVal Text As String = "Sub [If]") \' [If] note\nEnd Sub';
	expect(implementInterface({ source, moduleSources: { IJob: interfaceSource } })).toEqual({ ok: true, title: "Implement 1 member of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: '\nPrivate Sub [IJob_If](ByVal [If] As Long, Optional ByVal Text As String = "Sub [If]") \' [If] note\n    Err.Raise 5 \'TODO: implement this interface member\nEnd Sub\n' }] });
});
it('preserves type suffixes and continued headers with leading whitespace and default access', () => {
	const source = 'Implements IJob\r\n';
	const interfaceSource = '  Function Count%( _\r\n    ByVal Arg As Long)\r\nEnd Function';
	expect(implementInterface({ source, moduleSources: { IJob: interfaceSource } })).toEqual({ ok: true, title: "Implement 1 member of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: '\r\nPrivate Function IJob_Count%( _\r\n    ByVal Arg As Long)\r\n    Err.Raise 5 \'TODO: implement this interface member\r\nEnd Function\r\n' }] });
});

for (const count of [10, 1000]) it(`avoids ${count} dynamic declaration-name regular expressions`, () => {
	const source = 'Implements IJob\n';
	const interfaceSource = Array.from({ length: count }, (_, i) => `Public Sub Work${i}()\nEnd Sub\n`).join('');
	parseModule(source); parseModule(interfaceSource);
	const constructor = vi.spyOn(globalThis, 'RegExp');
	try {
		const result = implementInterface({ source, moduleSources: { IJob: interfaceSource } });
		expect(result.ok && result.title).toBe(`Implement ${count} members of 'IJob'`);
		expect(constructor.mock.calls.filter(([pattern]) => typeof pattern === 'string' && pattern.startsWith('(\\b(?:Sub|Function|Property'))).toHaveLength(0);
	} finally { constructor.mockRestore(); }
});
