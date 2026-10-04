// Procedure names in strings (issue #217): Application.Run, OnTime, OnAction,
// and a framework's wiring such as ReDim's `.OnClick "Demo.BuildReport"`.

import { describe, expect, it } from 'vitest';
import { macroNameCandidates, macroNameStringAt, macroNameTarget, resolveMacroNameCompletions } from '../src/analyzer/completion/macroNames';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../src/vbaProjectAnalysis';
import { resolveSignatureHelp, type SignatureHelpContext } from '../src/analyzer/signature/signatureHelp';

const BUTTON = 'Option Explicit\nPublic Function OnClick(ByVal handlerProc As String) As Btn\n    Set OnClick = Me\nEnd Function\n'
	+ 'Public Function Text(ByVal caption As String) As Btn\n    Set Text = Me\nEnd Function\n';
const DEMO = 'Option Explicit\nPublic Sub BuildReport()\nEnd Sub\nPublic Function Total() As Long\nEnd Function\n';

function contextFor(source: string): SignatureHelpContext {
	const modules = [
		{ moduleName: 'Main', source },
		{ moduleName: 'Demo', source: DEMO },
		{ moduleName: 'Btn', source: BUTTON, type: 'class' },
	];
	const project = buildVbaProjectIndex(modules);
	const procedures = projectProcedureSignatures(project);
	const options = projectAnalysisOptionsForModule(project, 'Main', procedures);
	// The editor hands the resolvers a flat list, as the project context does.
	return { ...options, projectProcedures: [...procedures.values()].flat(), moduleName: 'Main', moduleSource: source } as SignatureHelpContext;
}

function at(source: string, marker: string): number {
	return source.indexOf(marker) + marker.length;
}

describe('a string that names a procedure (issue #217)', () => {
    it.each([
        ['""', '', true],
        ['""""', '"', true],
        ['"a"""', 'a"', true],
        ['"a""', 'a"', false],
    ])('preserves quote boundaries for %s', (raw, text, closed) => {
        const source = 'Sub Demo()\nApplication.Run ' + raw + '\nEnd Sub';
        const start = source.indexOf(raw);
        const macro = macroNameStringAt(source, start + 1, contextFor(source));
        expect(macro?.text).toBe(text);
        expect(macro?.contentSpan).toEqual({ start: start + 1, end: start + raw.length - (closed ? 1 : 0) });
        if (closed) { expect(resolveMacroNameCompletions(source, start + raw.length, contextFor(source))).toBeUndefined(); }
    });

	it('is found where the parameter is named for one', () => {
		const wiring = 'Option Explicit\nSub Wire()\n    Dim b As New Btn\n    b.Text("Run").OnClick "Demo.Bu"\nEnd Sub\n';
		expect(macroNameStringAt(wiring, at(wiring, '"Demo.Bu'), contextFor(wiring))?.text).toBe('Demo.Bu');
		const run = 'Option Explicit\nSub Go()\n    Application.Run "Demo.BuildReport"\nEnd Sub\n';
		expect(macroNameStringAt(run, at(run, '"Demo'), contextFor(run))?.text).toBe('Demo.BuildReport');
		const onTime = 'Option Explicit\nSub Go()\n    Application.OnTime Now, "Demo.BuildReport"\nEnd Sub\n';
		expect(macroNameStringAt(onTime, at(onTime, '"Demo'), contextFor(onTime))?.text).toBe('Demo.BuildReport');
		const onAction = 'Option Explicit\nSub Go()\n    ActiveSheet.Shapes(1).OnAction = "Demo.BuildReport"\nEnd Sub\n';
		expect(macroNameStringAt(onAction, at(onAction, '"Demo'), contextFor(onAction))?.text).toBe('Demo.BuildReport');
		const named = 'Option Explicit\nSub Wire()\n    Dim b As New Btn\n    b.OnClick handlerProc:="Demo.BuildReport"\nEnd Sub\n';
		expect(macroNameStringAt(named, at(named, '"Demo'), contextFor(named))?.text).toBe('Demo.BuildReport');
	});

	it('is not found in any other string', () => {
		const caption = 'Option Explicit\nSub Wire()\n    Dim b As New Btn\n    b.Text("Demo.BuildReport").OnClick "x"\nEnd Sub\n';
		expect(macroNameStringAt(caption, at(caption, '"Demo'), contextFor(caption))).toBeUndefined();
		const plain = 'Option Explicit\nSub Go()\n    Debug.Print "Demo.BuildReport"\nEnd Sub\n';
		expect(macroNameStringAt(plain, at(plain, '"Demo'), contextFor(plain))).toBeUndefined();
	});

	it('completes inside the quotes and hovers the procedure named', () => {
		const wiring = 'Option Explicit\nSub Wire()\n    Dim b As New Btn\n    b.OnClick "Demo.Bu"\nEnd Sub\n';
		const ctx = contextFor(wiring);
		const completion = resolveMacroNameCompletions(wiring, at(wiring, '"Demo.Bu'), ctx);
		expect(completion?.contentSpan).toEqual({ start: wiring.indexOf('"Demo') + 1, end: wiring.indexOf('Bu"') + 2 });
		expect(completion?.candidates.map((candidate) => candidate.name)).toContain('Demo.BuildReport');
		expect(resolveMacroNameCompletions(wiring, at(wiring, 'Bu"'), ctx)).toBeUndefined();
		const full = 'Option Explicit\nSub Wire()\n    Dim b As New Btn\n    b.OnClick "Demo.BuildReport"\nEnd Sub\n';
		const hover = resolveHover(full, at(full, '"Demo.Bui'), contextFor(full));
		expect(hover?.signature).toContain('BuildReport');
		expect(hover?.details).toContain('Declared in Module: Demo');
	});

	it('gives the call tip of the last member of a chained statement call', () => {
		const chain = 'Option Explicit\nSub Wire()\n    Dim b As New Btn\n    b.Text("Run").OnClick "x"\nEnd Sub\n';
		expect(resolveSignatureHelp(chain, at(chain, 'OnClick "'), contextFor(chain))?.label).toContain('OnClick(');
	});

	it('offers the project procedures and finds the one named', () => {
		const ctx = contextFor('Option Explicit\n');
		expect(macroNameCandidates(ctx).map((candidate) => candidate.name)).toEqual(expect.arrayContaining(['Demo.BuildReport', 'Demo.Total']));
		expect(macroNameCandidates(ctx).some((candidate) => candidate.name.startsWith('Btn.'))).toBe(false);
		expect(macroNameTarget('Demo.BuildReport', ctx)?.name).toBe('BuildReport');
		expect(macroNameTarget('buildreport', ctx)?.moduleName).toBe('Demo');
		expect(macroNameTarget('Other.BuildReport', ctx)).toBeUndefined();
	});
});
