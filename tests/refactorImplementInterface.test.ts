// Implement Interface (issue #69). Signatures are copied from the interface's
// own text rather than rebuilt, because VBA rejects an implementing member
// whose signature does not match to the letter - a lost Optional or a ByRef
// turned ByVal and the class will not compile.

import { describe, expect, it } from 'vitest';
import { implementInterface } from '../src/analyzer/refactor/implementInterface';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

const IWIDGET = [
	'Option Explicit',
	'',
	'Public Sub Draw(ByVal x As Long, Optional ByVal y As Long = 0)',
	'End Sub',
	'',
	'Public Function Area() As Double',
	'End Function',
	'',
	'Public Property Get Name() As String',
	'End Property',
	'',
].join('\r\n');

function run(source: string, modules: Record<string, string> = { IWidget: IWIDGET }, interfaceName?: string) {
	return implementInterface({
		source,
		moduleSources: modules,
		...(interfaceName ? { interfaceName } : {}),
	});
}

function applied(source: string, modules?: Record<string, string>, interfaceName?: string): string {
	const result = run(source, modules, interfaceName);
	if (!result.ok) { throw new Error(`refused: ${result.reason}`); }
	return applyVbaTextEdits(source, result.edits);
}

function reason(source: string, modules?: Record<string, string>, interfaceName?: string): string {
	const result = run(source, modules, interfaceName);
	if (result.ok) { throw new Error(`expected a refusal, got ${result.title}`); }
	return result.reason;
}

describe('what it writes', () => {
	const KLASS = 'Option Explicit\r\nImplements IWidget\r\n';

	it('stubs every member, Private and prefixed with the interface name', () => {
		const out = applied(KLASS);
		expect(out).toContain('Private Sub IWidget_Draw(ByVal x As Long, Optional ByVal y As Long = 0)');
		expect(out).toContain('Private Function IWidget_Area() As Double');
		expect(out).toContain('Private Property Get IWidget_Name() As String');
	});

	it('copies the signature verbatim, Optional and default included', () => {
		// Rebuilding this is where the drift lives: VBA compares signatures
		// letter for letter.
		expect(applied(KLASS)).toContain('(ByVal x As Long, Optional ByVal y As Long = 0)');
	});

	it('raises rather than returning quietly', () => {
		const out = applied(KLASS);
		expect(out).toContain("    Err.Raise 5 'TODO: implement this interface member");
		expect(out.match(/Err\.Raise 5/g)).toHaveLength(3);
	});

	it('closes each stub with the right keyword', () => {
		const out = applied(KLASS);
		expect(out).toContain('End Sub');
		expect(out).toContain('End Function');
		expect(out).toContain('End Property');
	});

	it('writes only the members the class has not got', () => {
		const partial = [
			'Option Explicit',
			'Implements IWidget',
			'',
			'Private Function IWidget_Area() As Double',
			'End Function',
			'',
		].join('\r\n');
		const result = run(partial);
		expect(result.ok && result.title).toBe("Implement 2 members of 'IWidget'");
		const out = applied(partial);
		expect(out.match(/IWidget_Area/g)).toHaveLength(1);
		expect(out).toContain('IWidget_Draw');
		expect(out).toContain('IWidget_Name');
	});

	it('gives a public field of the interface its Get and Let pair', () => {
		const iface = 'Option Explicit\r\nPublic Total As Long\r\n';
		const out = applied('Implements IData\r\n', { IData: iface });
		expect(out).toContain('Private Property Get IData_Total() As Long');
		expect(out).toContain('Private Property Let IData_Total(ByVal RHS As Long)');
	});

	it('gives an object field Set rather than Let', () => {
		const iface = 'Option Explicit\r\nPublic Sheet As Worksheet\r\n';
		const out = applied('Implements IData\r\n', { IData: iface });
		expect(out).toContain('Private Property Set IData_Sheet(ByVal RHS As Worksheet)');
	});

	it('picks the named interface when the class implements two', () => {
		const source = 'Implements IWidget\r\nImplements IData\r\n';
		const modules = { IWidget: IWIDGET, IData: 'Public Sub Load()\r\nEnd Sub\r\n' };
		expect(applied(source, modules, 'IData')).toContain('Private Sub IData_Load()');
	});

	it('keeps the module\'s own line endings', () => {
		expect(applied('Implements IWidget\n')).not.toContain('\r');
	});
});

describe('what it refuses, and why', () => {
	it('refuses a class that implements nothing', () => {
		expect(reason('Option Explicit\r\n')).toMatch(/implements no interface/);
	});

	it('refuses an interface the class does not name', () => {
		expect(reason('Implements IWidget\r\n', { IWidget: IWIDGET, IOther: '' }, 'IOther'))
			.toMatch(/does not implement 'IOther'/);
	});

	it('refuses an interface the project has not got', () => {
		expect(reason('Implements IMissing\r\n', { IWidget: IWIDGET }))
			.toMatch(/no module called 'IMissing'/);
	});

	it('refuses an interface with no public members', () => {
		const empty = 'Option Explicit\r\nPrivate Sub Hidden()\r\nEnd Sub\r\n';
		expect(reason('Implements IEmpty\r\n', { IEmpty: empty })).toMatch(/no public members/);
	});

	it('refuses a class that has them all already', () => {
		const done = [
			'Implements IWidget',
			'Private Sub IWidget_Draw(ByVal x As Long, Optional ByVal y As Long = 0)',
			'End Sub',
			'Private Function IWidget_Area() As Double',
			'End Function',
			'Private Property Get IWidget_Name() As String',
			'End Property',
			'',
		].join('\r\n');
		expect(reason(done)).toMatch(/already implemented in full/);
	});

	it('asks which one when the class implements several and none is named', () => {
		const source = 'Implements IWidget\r\nImplements IData\r\n';
		expect(reason(source, { IWidget: IWIDGET, IData: '' })).toMatch(/Say which one/);
	});
});
