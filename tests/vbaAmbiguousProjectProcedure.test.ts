import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import type { VbaProcedureSignature } from '../src/analyzer/symbols/symbolModel';

// VBA is content for two modules to export the same public name; it refuses to
// compile an UNQUALIFIED call to it from a module declaring neither, with
// "Ambiguous name detected". The analyzer reported nothing, so a project the
// VBE will not compile read as 0 Errors.
//
// The finding belongs at the call site: exporting a name twice and always
// qualifying the calls is legal VBA, so flagging the declarations would fire on
// every legitimate project that does it.
function signature(name: string, moduleName: string): VbaProcedureSignature {
	return {
		name,
		moduleName,
		kind: 'sub',
		params: [{ name: 'label', type: 'String', optional: false, byRef: false }],
		visibility: 'Public',
	} as VbaProcedureSignature;
}

/** Helpers and Rival both export Recalculate, exactly as the issue describes. */
function twoOwners(): Map<string, VbaProcedureSignature[]> {
	return new Map([
		['recalculate', [signature('Recalculate', 'Helpers'), signature('Recalculate', 'Rival')]],
	]);
}

function analyze(moduleName: string, lines: string[], projectProcedures = twoOwners()) {
	return analyzeVbaModuleSource({
		source: lines.join('\r\n'),
		moduleName,
		moduleType: 'standard',
		moduleKind: 'standard',
		projectProcedures,
	}).diagnostics;
}

const codes = (ds: ReturnType<typeof analyze>) => ds.map((d) => String(d.code));

describe('ambiguous unqualified call to a name two modules export', () => {
	it('reports the bare call VBA refuses to compile', () => {
		const found = analyze('Consumer', [
			'Option Explicit',
			'',
			'Public Sub Drive()',
			'    Helpers.Recalculate "qualified"',
			'    Recalculate "ambiguous"',
			'End Sub',
			'',
		]);
		expect(codes(found)).toContain('ambiguous-project-procedure');
		const hit = found.find((d) => String(d.code) === 'ambiguous-project-procedure');
		expect(hit?.severity).toBe('error');
		// It must name both owners so the developer can pick one.
		expect(hit?.message).toContain('Helpers');
		expect(hit?.message).toContain('Rival');
		// Exactly one finding: the qualified call on the line above is fine.
		expect(found.filter((d) => String(d.code) === 'ambiguous-project-procedure')).toHaveLength(1);
	});

	it('stays silent on a qualified call', () => {
		expect(codes(analyze('Consumer', [
			'Option Explicit', '', 'Public Sub Drive()',
			'    Helpers.Recalculate "qualified"', 'End Sub', '',
		]))).not.toContain('ambiguous-project-procedure');
	});

	it('stays silent inside a module that declares the name itself', () => {
		// Module-local scope settles it; this is the issue's `Nearby` case.
		expect(codes(analyze('Helpers', [
			'Option Explicit',
			'',
			'Public Sub Recalculate(ByVal label As String)',
			'    Debug.Print label',
			'End Sub',
			'',
			'Public Sub Nearby()',
			'    Recalculate "same module"',
			'End Sub',
			'',
		]))).not.toContain('ambiguous-project-procedure');
	});

	it('stays silent when only one module exports the name', () => {
		const single = new Map([['recalculate', [signature('Recalculate', 'Helpers')]]]);
		expect(codes(analyze('Consumer', [
			'Option Explicit', '', 'Public Sub Drive()',
			'    Recalculate "only one"', 'End Sub', '',
		], single))).not.toContain('ambiguous-project-procedure');
	});

	it('stays silent when a local shadows the name', () => {
		expect(codes(analyze('Consumer', [
			'Option Explicit',
			'',
			'Public Sub Drive()',
			'    Dim Recalculate As String',
			'    Recalculate = "a local, not a call"',
			'End Sub',
			'',
		]))).not.toContain('ambiguous-project-procedure');
	});

	it('stays silent when the second declaration is Private', () => {
		// A Private procedure is not exported, so there is no collision.
		const onePrivate = new Map([[
			'recalculate',
			[
				signature('Recalculate', 'Helpers'),
				{ ...signature('Recalculate', 'Rival'), visibility: 'Private' } as VbaProcedureSignature,
			],
		]]);
		expect(codes(analyze('Consumer', [
			'Option Explicit', '', 'Public Sub Drive()',
			'    Recalculate "one is private"', 'End Sub', '',
		], onePrivate))).not.toContain('ambiguous-project-procedure');
	});

	it('stays silent with no project context at all', () => {
		// Built without analyze(): passing undefined would trigger its default.
		const found = analyzeVbaModuleSource({
			source: ['Option Explicit', '', 'Public Sub Drive()',
				'    Recalculate "no project map"', 'End Sub', ''].join('\r\n'),
			moduleName: 'Consumer',
			moduleType: 'standard',
			moduleKind: 'standard',
		}).diagnostics;
		expect(found.map((d) => String(d.code))).not.toContain('ambiguous-project-procedure');
	});
});
