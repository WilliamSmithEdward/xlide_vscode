// A call with too few arguments. Measured in Excel, Word, PowerPoint and
// Access 16.0 (build 20527) through the VBE oracle: the VBE says "Argument not
// optional", for a project procedure, a VBA function, `VBA.Left` or a host
// method alike, except where it parses the call as syntax of its own and says
// "Syntax error": a bare call of Len, LenB, InStr, InStrB, StrComp, String$,
// the C conversions but CDec and CVDate, Abs, Int, Fix, Sgn, LBound or
// UBound, a Mid statement's target and Debug.Assert. One too many is "Wrong
// number of arguments or invalid property assignment". The cases are the
// too_few_arguments_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../../src/vbaProjectAnalysis';
import { oracleDefaultDocumentModules, oracleModuleInput, oracleProjectReferences, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	description: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: OracleFixtureModule[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('too_few_arguments_'));

/** The argument-count messages, for the project as the extension reads it. */
function argumentMessages(kase: { host: string; modules: OracleFixtureModule[] }): string[] {
	const modules = kase.modules.map(oracleModuleInput);
	const all = [...modules, ...oracleDefaultDocumentModules(kase, kase.modules).map(oracleModuleInput)];
	const project = buildVbaProjectIndex(all, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: kase.host,
		...oracleProjectReferences({ host: kase.host }),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.code === 'argument-count').map((diagnostic) => diagnostic.message));
}

/** Every error code, for the project as the extension reads it. */
function codes(kase: { host: string; modules: OracleFixtureModule[] }): string[] {
	const modules = kase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: kase.host,
		...oracleProjectReferences({ host: kase.host }), ...projectAnalysisOptionsForModule(project, mod.moduleName, projectProcedureSignatures(project)),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? ''));
}

const procedure = (...lines: string[]) => ({
	host: 'excel',
	modules: [{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim v As Variant, s As String', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n') }],
});

describe('too few arguments against the measured fixtures', () => {
	it('has fixtures for every Office host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s gets the VBE message', (_id, fixture) => {
		const vbe = /Compile error: (.+?)\. Measured/.exec(fixture.description)![1];
		const messages = argumentMessages(fixture);
		expect(messages).toHaveLength(1);
		if (vbe.startsWith('Wrong number of arguments')) {
			expect(messages[0]).toMatch(/^Wrong number of arguments to /);
		} else {
			expect(messages[0]).toMatch(new RegExp(`This is a VBE compile error: ${vbe}\\.$`));
		}
	});
});

describe('too few arguments elsewhere', () => {
	it('names the parameter left out', () => {
		expect(argumentMessages(procedure('s = Left("a")'))).toEqual(["'Length' is required by 'Left', which got 1 argument. This is a VBE compile error: Argument not optional."]);
	});

	it('reads VBA.CLng() as an ordinary call, not syntax of its own', () => {
		expect(argumentMessages(procedure('v = VBA.CLng()'))[0]).toMatch(/Argument not optional\.$/);
		expect(argumentMessages(procedure('v = CLng()'))[0]).toMatch(/Syntax error\.$/);
	});

	it('reads Mid as a function with one argument, and Mid as a statement target', () => {
		expect(argumentMessages(procedure('v = Mid("a")'))[0]).toMatch(/Argument not optional\.$/);
		expect(argumentMessages(procedure('Mid(s) = "x"'))[0]).toMatch(/Syntax error\.$/);
		expect(argumentMessages(procedure('If v Then Mid$(s) = "x"'))[0]).toMatch(/Syntax error\.$/);
	});

	it('leaves an explicit Call without parentheses to its own rule, whatever the spacing', () => {
		expect(argumentMessages(procedure('Call MsgBox "hello"'))).toEqual([]);
		expect(argumentMessages(procedure('Call MsgBox  "hello"'))).toEqual([]);
		expect(argumentMessages(procedure('Call MsgBox()'))[0]).toMatch(/Argument not optional\.$/);
		// Else ends the branch: the Call has no argument at all.
		expect(argumentMessages(procedure('If True Then Call MsgBox Else v = 1'))[0]).toMatch(/Argument not optional\.$/);
		// In a one-line If's branch call-requires-parens does not run, so it is reported here, as the VBE's Syntax error.
		const branchCall = "The Call to 'MsgBox' gives its arguments outside parentheses. This is a VBE compile error: Syntax error.";
		expect(argumentMessages(procedure('If v Then Call MsgBox "hello"'))).toEqual([branchCall]);
		expect(argumentMessages(procedure('If v Then v = 1 Else Call MsgBox "hello"'))).toEqual([branchCall]);
	});

	it('leaves the statement form to call-requires-parens, which reports it', () => {
		expect(codes(procedure('Call MsgBox "hello"'))).toEqual(['call-requires-parens']);
	});

	it('pins the Syntax error wording', () => {
		expect(argumentMessages(procedure('v = CInt()'))).toEqual(["The call to 'CInt' gives 0 of its 1 required argument. This is a VBE compile error: Syntax error."]);
	});

	it('reads a bracketed name as an ordinary call', () => {
		expect(argumentMessages(procedure('v = [Len]()'))[0]).toMatch(/Argument not optional\.$/);
	});

	it('reads a Mid statement past a line number, a continuation and nested calls, and not a comparison', () => {
		expect(argumentMessages(procedure('10 Mid(s) = "x"'))[0]).toMatch(/Syntax error\.$/);
		expect(argumentMessages(procedure('Mid _', '    (s) = "x"'))[0]).toMatch(/Syntax error\.$/);
		expect(argumentMessages(procedure('Mid(IIf(True, CStr(Len(s)), s)) = "x"'))[0]).toMatch(/Syntax error\.$/);
		expect(argumentMessages(procedure('If Mid(s) = "a" Then v = 1'))[0]).toMatch(/Argument not optional\.$/);
	});

	it("reads Debug.Assert as VBA's, past any spacing, and a project member reached through .Debug by its own signature", () => {
		expect(argumentMessages(procedure('Debug.   Assert'))[0]).toMatch(/Syntax error\.$/);
		const kase = { host: 'excel', modules: [
			{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim h As New Holder', '    h.Debug.Assert', 'End Sub', ''].join('\r\n') },
			{ name: 'Holder', type: 'class', source: ['Option Explicit', 'Public Function Debug() As Thing', 'End Function', ''].join('\r\n') },
			{ name: 'Thing', type: 'class', source: ['Option Explicit', 'Public Sub Assert(ByVal a As Long)', 'End Sub', ''].join('\r\n') },
		] };
		expect(argumentMessages(kase)[0]).toMatch(/Argument not optional\.$/);
	});

	it('classifies a named call that leaves out a required argument the same way', () => {
		expect(argumentMessages(procedure('v = LBound(Dimension:=1)'))[0]).toMatch(/Syntax error\.$/);
		expect(argumentMessages(procedure('v = Left(Length:=1)'))[0]).toMatch(/Argument not optional\.$/);
	});

	it('reads a project function named like a VBA one by its own signature', () => {
		const kase = { host: 'excel', modules: [
			{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Function Abs(a As Long, b As Long) As Long', 'End Function', 'Sub Main()', '    Dim v As Variant', '    v = Abs(1)', 'End Sub', ''].join('\r\n') },
		] };
		expect(argumentMessages(kase)[0]).toMatch(/^'b' is required by 'Abs'.*Argument not optional\.$/);
	});
});
