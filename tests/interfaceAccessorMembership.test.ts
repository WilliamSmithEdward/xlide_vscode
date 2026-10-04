import { expect, it, vi } from 'vitest';
import { implementInterface } from '../src/analyzer/refactor/implementInterface';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';

const pairs = [
	{ type: 'Long', setter: 'Let' },
	{ type: 'Object', setter: 'Set' },
] as const;
for (const { type, setter } of pairs) for (const declaration of ['field', 'properties']) for (const present of ['Get', setter]) for (const eol of ['\n', '\r\n', '\r']) {
	it(`adds missing ${present === 'Get' ? setter : 'Get'} for ${declaration} ${type} with ${JSON.stringify(eol)}`, () => {
		const getter = `Property Get Value() As ${type}`;
		const writer = `Property ${setter} Value(ByVal RHS As ${type})`;
		const interfaceSource = declaration === 'field' ? `Public Value As ${type}` : ['Public ' + getter, 'End Property', 'Public ' + writer, 'End Property'].join(eol);
		const existing = present === 'Get' ? getter : writer;
		const missing = present === 'Get' ? writer : getter;
		const source = ['Implements IData', 'Private ' + existing.replace('Value', 'idata_value'), 'End Property', ''].join(eol);
		const generatedEol = eol;
		const text = (source.endsWith(generatedEol) ? '' : generatedEol) + generatedEol + 'Private ' + missing.replace('Value', 'IData_Value') + generatedEol + "    Err.Raise 5 'TODO: implement this interface member" + generatedEol + 'End Property' + generatedEol;
		const expected = { ok: true, title: "Implement 1 member of 'IData'", edits: [{ span: { start: source.length, end: source.length }, newText: text }] };
		const result = implementInterface({ source, moduleSources: { idata: interfaceSource } });
		expect(result).toEqual(expected);
		if (!result.ok) { throw new Error(result.reason); }
		const procedures = parseModule(applyVbaTextEdits(source, result.edits)).members.filter(member => member.kind === 'Procedure');
		expect(procedures.map(member => member.procKind).sort()).toEqual(['PropertyGet', `Property${setter}`].sort());
	});
}
for (const { type, setter } of pairs) for (const eol of ['\n', '\r\n', '\r']) {
	it(`recognizes both ${type} field accessors with ${JSON.stringify(eol)}`, () => {
		const source = ['Implements IData', `Private Property Get IData_Value() As ${type}`, 'End Property', `Private Property ${setter} IData_Value(ByVal RHS As ${type})`, 'End Property'].join(eol);
		expect(implementInterface({ source, moduleSources: { IData: `Public Value As ${type}` } })).toEqual({ ok: false, reason: "'IData' is already implemented in full." });
	});
}
for (const [wanted, existing] of [
	['Function Work() As Long', 'Sub IData_Work()'],
	['Sub Work()', 'Function IData_Work() As Long'],
	['Property Get Work() As Long', 'Function IData_Work() As Long'],
	['Sub Work()', 'Property Get IData_Work() As Long'],
]) {
	it(`refuses an incompatible existing ${existing}`, () => {
		const closer = (header: string) => header.startsWith('Sub') ? 'Sub' : header.startsWith('Function') ? 'Function' : 'Property';
		const source = 'Implements IData\nPrivate ' + existing + '\nEnd ' + closer(existing);
		const interfaceSource = 'Public ' + wanted + '\nEnd ' + closer(wanted);
		expect(implementInterface({ source, moduleSources: { IData: interfaceSource } })).toEqual({ ok: false, reason: "The class already has 'IData_Work' as a different procedure kind. Rename or correct it before implementing 'IData'." });
	});
}
it('does not let a different property writer satisfy a required setter', () => {
	const source = 'Implements IData\nPrivate Property Get IData_Value() As Object\nEnd Property\nPrivate Property Let IData_Value(ByVal RHS As Variant)\nEnd Property\n';
	const result = implementInterface({ source, moduleSources: { IData: 'Public Value As Object' } });
	expect(result.ok && result.title).toBe("Implement 1 member of 'IData'");
	expect(result.ok && result.edits[0].newText).toBe("\nPrivate Property Set IData_Value(ByVal RHS As Object)\n    Err.Raise 5 'TODO: implement this interface member\nEnd Property\n");
});

for (const count of [10, 1000]) it(`does not index ${count} unrelated procedure names`, () => {
	const source = 'Implements IData\n' + Array.from({ length: count }, (_, i) => `Private Sub Noise${i}()\nEnd Sub\n`).join('');
	const interfaceSource = 'Public Value As Long';
	parseModule(source); parseModule(interfaceSource);
	const add = vi.spyOn(Set.prototype, 'add');
	try {
		const result = implementInterface({ source, moduleSources: { IData: interfaceSource } });
		expect(result.ok && result.title).toBe("Implement 2 members of 'IData'");
		expect(add.mock.calls.filter(([value]) => typeof value === 'string' && /^noise\d+$/.test(value))).toHaveLength(0);
	} finally { add.mockRestore(); }
});
