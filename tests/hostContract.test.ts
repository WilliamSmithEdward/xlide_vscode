import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import { moduleKindFromType } from '../src/vbaProjectAnalysis';
import { hostObjectModelForToken, EMPTY_HOST_MODEL } from '../src/analyzer/host/hostRegistry';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';

// Guards the claims in docs/host_contract.md. That document tells a host
// outside this repo - the VBE add-in, whose CodeModule returns code and nothing
// else - which facts it must supply and what happens when it does not. Every
// row of it is asserted here, so the document cannot quietly stop being true.

const NL = '\n';
const EMPTY = new Set<string>();

function codes(source: string, options: Record<string, unknown>): string[] {
	return analyzeModule(source, { moduleName: 'M', knownIdentifiers: EMPTY, ...options } as never)
		.map((hit) => hit.code)
		.sort();
}

describe('moduleKind - required in practice', () => {
	it('maps the four kinds', () => {
		expect(moduleKindFromType('standard')).toBe('standard');
		expect(moduleKindFromType('class')).toBe('class');
		expect(moduleKindFromType('document')).toBe('document');
		expect(moduleKindFromType('userform')).toBe('userform');
	});

	it('falls back to standard for absent AND for an unrecognised string', () => {
		// Documented as a FAIL-SILENT: there is no error path, so a host typo
		// degrades rather than complains. The doc warns about it; this pins it
		// so the warning stays accurate.
		expect(moduleKindFromType(undefined)).toBe('standard');
		expect(moduleKindFromType('Class')).toBe('standard');
		expect(moduleKindFromType('vbext_ct_ClassModule')).toBe('standard');
	});

	it('is what decides whether object-module rules run', () => {
		const me = ['Option Explicit', 'Public Sub P()', '    Debug.Print Me.Name', 'End Sub', ''].join(NL);
		expect(codes(me, { moduleKind: 'class' })).toEqual([]);
		expect(codes(me, { moduleKind: 'document' })).toEqual([]);
		expect(codes(me, { moduleKind: 'userform' })).toEqual([]);
		// The cost of getting the kind wrong: a valid class reports.
		expect(codes(me, { moduleKind: 'standard' })).toEqual(['me-outside-object-module']);
		expect(codes(me, {})).toEqual(['me-outside-object-module']);
	});
});

describe('documentType - inferred from the code name when absent', () => {
	const HANDLERS = {
		workbook: 'Private Sub Workbook_Open()' + NL + 'End Sub',
		worksheet: 'Private Sub Worksheet_Change(ByVal Target As Range)' + NL + 'End Sub',
	};

	function flagged(moduleName: string, handler: string): boolean {
		const source = ['Option Explicit', handler, ''].join(NL);
		return codes(source, { moduleName, moduleKind: 'document' })
			.includes('event-handler-module-scope');
	}

	it('reads the conventional code names correctly with no host answer', () => {
		expect(flagged('ThisWorkbook', HANDLERS.workbook), 'ThisWorkbook/Workbook_Open').toBe(false);
		expect(flagged('Sheet1', HANDLERS.worksheet), 'Sheet1/Worksheet_Change').toBe(false);
		expect(flagged('Sheet12', HANDLERS.worksheet), 'Sheet12/Worksheet_Change').toBe(false);
		// Anything unrecognised defaults to worksheet, which is right for a
		// RENAMED sheet code name - the common case.
		expect(flagged('SalesData', HANDLERS.worksheet), 'SalesData/Worksheet_Change').toBe(false);
	});

	it('still catches a handler in the wrong document module', () => {
		expect(flagged('Sheet1', HANDLERS.workbook), 'Workbook_Open in a sheet').toBe(true);
		expect(flagged('ThisWorkbook', HANDLERS.worksheet), 'Worksheet_Change in the book').toBe(true);
	});

	it('is overridden by the host answer when supplied', () => {
		const source = ['Option Explicit', HANDLERS.workbook, ''].join(NL);
		// The name says worksheet, the host says workbook, and the host wins.
		expect(codes(source, { moduleName: 'Dashboard', moduleKind: 'document', documentType: 'workbook' }))
			.toEqual([]);
	});
});

describe('conditionalCompilation - decides which branch is analyzed', () => {
	const source = [
		'Option Explicit', 'Public Sub P()',
		'#If VBA7 Then', '    OnlyInVba7 = 1',
		'#Else', '    OnlyInVba6 = 2',
		'#End If', 'End Sub', '',
	].join(NL);

	function undeclaredNames(env: unknown): string[] {
		return analyzeModule(source, {
			moduleName: 'M', knownIdentifiers: EMPTY, conditionalCompilation: env,
		} as never)
			.filter((hit) => hit.code === 'undeclared-variable')
			.map((hit) => /'([^']+)'/.exec(hit.message)?.[1] ?? '?');
	}

	it('takes the nested compilerConstants shape, not a bare record', () => {
		expect(undeclaredNames({ compilerConstants: { VBA7: false } })).toEqual(['OnlyInVba6']);
		expect(undeclaredNames({ compilerConstants: { VBA7: true } })).toEqual(['OnlyInVba7']);
		// A bare record is silently ignored - it sets no constant at all, so the
		// built-in defaults still decide. Worth pinning: it looks like it works.
		expect(undeclaredNames({ VBA7: false })).toEqual(['OnlyInVba7']);
	});

	it('defaults to the VBA7 branch when absent', () => {
		expect(undeclaredNames(undefined)).toEqual(['OnlyInVba7']);
	});
});

describe('host - three outcomes, and the middle one matters', () => {
	it('answers undefined for absent and for excel, so callers keep the Excel default', () => {
		expect(hostObjectModelForToken(undefined)).toBeUndefined();
		expect(hostObjectModelForToken('excel')).toBeUndefined();
		expect(hostObjectModelForToken('EXCEL')).toBeUndefined();
	});

	it('answers a named host its own model', () => {
		expect(hostObjectModelForToken('word')?.hostName).toBe(getWordObjectModel().hostName);
	});

	it('answers a host with no model the empty one, which asserts nothing', () => {
		expect(hostObjectModelForToken('visio')).toBe(EMPTY_HOST_MODEL);
		expect(EMPTY_HOST_MODEL.hostName).toBeUndefined();
		expect(Object.keys(EMPTY_HOST_MODEL.types)).toEqual([]);
	});
});

describe('source - offsets are measured against exactly the string supplied', () => {
	it('does not shift when a fact is supplied out of band', () => {
		// The reason predeclaredId is a FIELD and not a synthesized header line:
		// prepending anything to the source moves every span by its length.
		const source = ['Option Explicit', 'Public Sub P()', '    nope = 1', 'End Sub', ''].join(NL);
		const hit = analyzeModule(source, { moduleName: 'M', knownIdentifiers: EMPTY })
			.find((d) => d.code === 'undeclared-variable');
		expect(source.slice(hit?.span.start, hit?.span.end)).toBe('nope');
		expect(hit?.span.start).toBe(source.indexOf('nope'));
	});
});
