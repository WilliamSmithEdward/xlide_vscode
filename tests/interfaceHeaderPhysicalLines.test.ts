import { expect, it } from 'vitest';
import { implementInterface } from '../src/analyzer/refactor/implementInterface';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';

for (const eol of ['\n', '\r\n', '\r']) for (const bodyLines of [1, 1000]) {
	it(`copies only the interface header with ${bodyLines} body lines and ${JSON.stringify(eol)}`, () => {
		const source = 'Implements IJob\n';
		const interfaceSource = ['Public Sub Work(ByVal Input As Long)', ...Array.from({ length: bodyLines }, () => '    Debug.Print "private body"'), 'End Sub'].join(eol);
		const result = implementInterface({ source, moduleSources: { IJob: interfaceSource } });
		expect(result).toEqual({ ok: true, title: "Implement 1 member of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: "\nPrivate Sub IJob_Work(ByVal Input As Long)\n    Err.Raise 5 'TODO: implement this interface member\nEnd Sub\n" }] });
		if (!result.ok) { throw new Error(result.reason); }
		const applied = applyVbaTextEdits(source, result.edits);
		expect(parseModule(applied).members.filter(member => member.kind === 'Procedure')).toHaveLength(1);
		expect(applied).not.toContain('private body');
	});
}
for (const endings of [['\r', '\r'], ['\r\n', '\r'], ['\r', '\n'], ['\n', '\r\n']]) {
	it(`copies continued signatures with physical breaks ${JSON.stringify(endings)}`, () => {
		const source = 'Implements IJob\r\n';
		const header = 'Public Function Work(ByRef Input As Long, _' + endings[0] + '    Optional ByVal Extra As Long = 2) As String';
		const interfaceSource = header + endings[1] + '    Work = "private body"' + endings[1] + 'End Function';
		expect(implementInterface({ source, moduleSources: { IJob: interfaceSource } })).toEqual({ ok: true, title: "Implement 1 member of 'IJob'", edits: [{ span: { start: source.length, end: source.length }, newText: '\r\nPrivate ' + header.replace('Public ', '').replace('Function Work', 'Function IJob_Work') + "\r\n    Err.Raise 5 'TODO: implement this interface member\r\nEnd Function\r\n" }] });
	});
}
it('keeps LF/CRLF comments, access modifiers and multiple continued lines verbatim', () => {
	for (const eol of ['\n', '\r\n']) {
		const source = 'Implements IJob' + eol;
		const header = ['Friend Function Work( _', '    ByRef Input As Long, _', '    Optional ByVal Extra As Long = 2) As String \' header note'].join(eol);
		const result = implementInterface({ source, moduleSources: { IJob: header + eol + 'End Function' } });
		expect(result.ok && result.edits[0].newText).toBe(eol + 'Private ' + header.replace('Friend ', '').replace('Function Work', 'Function IJob_Work') + eol + "    Err.Raise 5 'TODO: implement this interface member" + eol + 'End Function' + eol);
	}
});
