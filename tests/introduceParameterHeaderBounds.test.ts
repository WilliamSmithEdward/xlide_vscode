import { expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for (const eol of ['\n', '\r\n', '\r']) for (const fixture of [
	{ header: 'Public Sub Report', after: 'Public Sub Report(ByVal limit As Long)', body: 'Debug.Print Len(")") + limit', caller: 'Report', call: 'Report 3' },
	{ header: "Public Sub Report ' header )", after: "Public Sub Report(ByVal limit As Long) ' header )", body: 'Debug.Print limit', caller: 'Report', call: 'Report 3' },
	{ header: 'Public Sub Report(ByRef items() As Long)', after: 'Public Sub Report(ByRef items() As Long, ByVal limit As Long)', body: 'Debug.Print limit', caller: 'Report values', call: 'Report values, 3' },
	{ header: ['Public Sub Report( _', 'ByRef items() As Long _', ')'].join(eol), after: ['Public Sub Report( _', 'ByRef items() As Long _', ', ByVal limit As Long)'].join(eol), body: 'Debug.Print limit', caller: 'Report values', call: 'Report values, 3' },
	{ header: 'Public Sub Report()', after: 'Public Sub Report(ByVal limit As Long)', body: 'Debug.Print limit', caller: 'Report', call: 'Report 3' },
]) it(`inserts in the actual header ${JSON.stringify(fixture.header)} with ${JSON.stringify(eol)}`, () => {
	const source = [fixture.header, 'Dim limit As Long', 'limit = 3', fixture.body, 'End Sub', 'Sub Caller()', 'Dim values() As Long', fixture.caller, 'End Sub', ''].join(eol);
	const result = introduceParameter({ source, offset: source.indexOf('limit As'), moduleName: 'M' });
	if (!result.ok) throw new Error(result.reason);
	expect(applyVbaTextEdits(source, result.edits)).toBe([fixture.after, fixture.body, 'End Sub', 'Sub Caller()', 'Dim values() As Long', fixture.call, 'End Sub', ''].join(eol));
});
