// sourceNameScopeFor's callableShadows answers has() from two layers: the
// procedure's own non-callable names over the module's. Each layer is pinned
// here directly, because the rules that read it have other checks that hide
// either layer going missing.

import { describe, expect, it } from 'vitest';
import type { ProcedureNode } from '../src/analyzer/parser/nodes';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { sourceNameScopeFor } from '../src/analyzer/diagnostics/typeInference';

const source = [
	'Private Total As Long',
	'Const Limit As Long = 10',
	'Sub Main(ByVal Count As Long)',
	'Dim Item As Long',
	'End Sub',
	'Sub Other()',
	'End Sub',
].join('\n');

function callableShadows(procName: string) {
	const module = parseModule(source);
	const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule: module });
	const proc = module.members.find(
		(member): member is ProcedureNode => member.kind === 'Procedure' && member.name === procName,
	);
	if (!proc) {
		throw new Error(`no procedure ${procName}`);
	}
	return sourceNameScopeFor(symbols, proc).callableShadows;
}

describe('sourceNameScopeFor callableShadows', () => {
	it("holds the procedure's own parameters and locals", () => {
		const shadows = callableShadows('Main');
		expect(shadows.has('count')).toBe(true);
		expect(shadows.has('item')).toBe(true);
	});

	it("holds the module's variables and constants under every procedure", () => {
		for (const procName of ['Main', 'Other']) {
			const shadows = callableShadows(procName);
			expect(shadows.has('total')).toBe(true);
			expect(shadows.has('limit')).toBe(true);
		}
	});

	it("keeps one procedure's names out of another's scope", () => {
		const shadows = callableShadows('Other');
		expect(shadows.has('count')).toBe(false);
		expect(shadows.has('item')).toBe(false);
	});

	it('does not hold procedure names or undeclared names', () => {
		const shadows = callableShadows('Main');
		expect(shadows.has('main')).toBe(false);
		expect(shadows.has('other')).toBe(false);
		expect(shadows.has('missing')).toBe(false);
	});
});
