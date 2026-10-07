import { afterEach, expect, it, vi } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { collectProcedureLabelDeclarations, collectProcedureLabelReferences } from '../src/analyzer/flow/procedureLabels';
import * as helpers from '../src/analyzer/lexer/tokenHelpers';
import type { ProcedureNode } from '../src/analyzer/parser/nodes';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
afterEach(() => vi.restoreAllMocks());
for (const count of [10, 100, 1000]) it(`reuses label scans across ten readers of ${count} statements`, () => {
	const source = 'Sub P()\nentry:\n' + 'Debug.Print 1\n'.repeat(count) + 'GoTo entry\nEnd Sub\n';
	const proc = parseModule(source).members[0] as ProcedureNode;
	const spy = vi.spyOn(helpers, 'statementTokensCached');
	for (let i = 0; i < 10; i++) {
		expect(collectProcedureLabelDeclarations(source, proc).map(label => label.text)).toEqual(['entry']);
		expect(collectProcedureLabelReferences(source, proc).map(label => label.text)).toEqual(['entry']);
	}
	expect(spy.mock.calls.length).toBeLessThanOrEqual(2 * proc.body.length);
});
it('isolates returned arrays, labels and spans', () => {
	const source = 'Sub P()\nentry:\nGoTo entry\nEnd Sub\n';
	const proc = parseModule(source).members[0] as ProcedureNode;
	const declarations = collectProcedureLabelDeclarations(source, proc), refs = collectProcedureLabelReferences(source, proc);
	const expectedDeclarations = structuredClone(declarations), expectedRefs = structuredClone(refs);
	declarations[0].key = 'changed'; declarations[0].span.start = -1; declarations.length = 0;
	refs[0].key = 'changed'; refs[0].span.end = -1; refs.push(refs[0]);
	expect(collectProcedureLabelDeclarations(source, proc)).toEqual(expectedDeclarations);
	expect(collectProcedureLabelReferences(source, proc)).toEqual(expectedRefs);
});
it('replaces facts for changed sources and conditional activity on a reused procedure', () => {
	const source = 'Sub P()\n#If FLAG Then\nfirst:\nGoTo first\n#Else\nother:\nGoTo other\n#End If\nEnd Sub\n';
	const module = parseModule(source), proc = module.members[0] as ProcedureNode;
	for (const flag of [true, false, true]) {
		const activity = createConditionalActivityTracker(module, { compilerConstants: { FLAG: flag } });
		expect(collectProcedureLabelDeclarations(source, proc, activity).map(label => label.text)).toEqual([flag ? 'first' : 'other']);
		expect(collectProcedureLabelReferences(source, proc, activity).map(label => label.text)).toEqual([flag ? 'first' : 'other']);
	}
	for (const text of [source, source.replaceAll('first', 'newer'), source]) {
		expect(collectProcedureLabelDeclarations(text, proc).map(label => label.text)).toEqual([text.includes('newer') ? 'newer' : 'first', 'other']);
	}
});
