import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import { getExcelObjectModel } from '../src/analyzer/host/excelObjectModel';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import { EMPTY_HOST_MODEL } from '../src/analyzer/host/hostRegistry';
import type { HostObjectModel } from '../src/analyzer/host/excelObjectModel';

// Two defects the undeclared-variable scanner had around the two characters
// VBA uses for host lookups.
//
// `!` is the DEFAULT-MEMBER accessor, so `rs!CustomerName` is member access
// exactly as `rs.CustomerName` is. The scanner skipped a name after `.` but not
// after `!`, so every field read through a recordset or an Access form was
// called an undefined variable. Host-neutral: this was never a variable.
//
// `[name]` is host-dependent. In Excel the brackets are shorthand for
// `Application.Evaluate`, so `[A1]` compiles and needs no declaration. Word has
// no such feature - measured in the VBE, `v = [foo]` with nothing declaring
// `foo` is a compile error - so the report is correct there and stays.

const NL = '\n';

function undeclared(body: readonly string[], hostModel?: HostObjectModel): string[] {
	const source = ['Option Explicit', '', 'Public Sub P()', ...body, 'End Sub', ''].join(NL);
	return analyzeModule(source, {
		moduleName: 'M',
		knownIdentifiers: new Set<string>(),
		hostModel,
	})
		.filter((hit) => hit.code === 'undeclared-variable')
		.map((hit) => /'([^']+)'/.exec(hit.message)?.[1] ?? '?');
}

describe('the bang operator is member access', () => {
	const HOSTS: ReadonlyArray<readonly [string, HostObjectModel | undefined]> = [
		['default', undefined],
		['Excel', getExcelObjectModel()],
		['Word', getWordObjectModel()],
		['unknown', EMPTY_HOST_MODEL],
	];

	it('does not call a banged member an undeclared variable, in any host', () => {
		for (const [label, model] of HOSTS) {
			expect(
				undeclared(['    Dim rs As Object', '    Debug.Print rs!Field1'], model),
				`bang in ${label}`,
			).toEqual([]);
			expect(
				undeclared(['    Dim rs As Object', '    Debug.Print rs![My Field]'], model),
				`bracketed bang in ${label}`,
			).toEqual([]);
		}
	});

	it('still reports the RECEIVER, which is a real name', () => {
		// Only the members are skipped. `Forms` has to resolve on its own.
		expect(undeclared(['    Dim v As Variant', '    v = Forms!frmMain!txtName']))
			.toEqual(['Forms']);
	});

	it('resolves the Access form idiom whole, under Access', () => {
		const access = getAccessObjectModel();
		expect(undeclared(['    Dim v As Variant', '    v = Forms!frmMain!txtName'], access))
			.toEqual([]);
		expect(undeclared(['    Dim v As Variant', '    v = Reports!rptA!ctl'], access))
			.toEqual([]);
	});

	it('leaves the Single type suffix alone', () => {
		// `!` is also the Single suffix. A suffix is only ever followed by an
		// operator or the end of the statement, never by a name, so the member
		// skip cannot swallow one.
		expect(undeclared(['    Dim x!', '    x! = 1'])).toEqual([]);
		expect(undeclared([
			'    Dim a!',
			'    Dim b As Long',
			'    a! = 1',
			'    b = a! + 2',
		])).toEqual([]);
	});
});

describe('a bracketed name on its own', () => {
	const READS: Record<string, readonly string[]> = {
		'read': ['    Dim v As Variant', '    v = [A1]'],
		'assignment': ['    [A1] = 5'],
		'member access on the result': ['    Dim v As Variant', '    v = [A1].Value'],
	};

	it('is an Evaluate lookup in Excel, so nothing is reported', () => {
		for (const [label, body] of Object.entries(READS)) {
			expect(undeclared(body, getExcelObjectModel()), `Excel ${label}`).toEqual([]);
		}
	});

	it('is reported in Word, which has no such shorthand', () => {
		for (const [label, body] of Object.entries(READS)) {
			expect(undeclared(body, getWordObjectModel()), `Word ${label}`).toEqual(['A1']);
		}
	});

	it('is reported in Access, which has no such shorthand either', () => {
		expect(undeclared(['    Dim v As Variant', '    v = [foo]'], getAccessObjectModel()))
			.toEqual(['foo']);
	});

	it('stays silent when the host is unstated or unknown', () => {
		// An absent model is Excel's by default (#28), and a model that knows
		// nothing asserts nothing. Neither may guess its way into a report.
		expect(undeclared(READS.read, undefined), 'absent').toEqual([]);
		expect(undeclared(READS.read, EMPTY_HOST_MODEL), 'empty').toEqual([]);
	});

	it('is fine everywhere once it is declared', () => {
		for (const model of [undefined, getExcelObjectModel(), getWordObjectModel()]) {
			expect(undeclared(['    Dim [My Var] As Long', '    [My Var] = 1'], model)).toEqual([]);
		}
	});

	it('does not stop a plain undeclared name being reported', () => {
		for (const model of [undefined, getExcelObjectModel(), getWordObjectModel()]) {
			expect(undeclared(['    Dim v As Variant', '    v = nope'], model)).toEqual(['nope']);
		}
	});
});
