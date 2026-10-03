import { describe, expect, it } from 'vitest';
import { procedureHasUnstructuredFlow } from '../src/analyzer/flow/procedureUnstructured';
import { collectProcedureLabelDeclarations, collectProcedureLabelReferences } from '../src/analyzer/flow/procedureLabels';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';

function procedure(source: string) {
	const module = parseModule(source);
	const proc = module.members.find(member => member.kind === 'Procedure');
	if (proc?.kind !== 'Procedure') { throw new Error('Missing procedure'); }
	return {module, proc};
}

describe('unstructured flow existence walk', () => {
	it.each(['GoTo Done', 'Entry:', 'On Error Resume Next', 'Resume Next'])('stops after an early witness: %s', (first) => {
		const source = 'Sub P()\n' + first + '\n' + 'x = x + 1\n'.repeat(200) + 'End Sub';
		const {proc} = procedure(source);
		let reads = 0;
		const watched = {...proc, body: proc.body.map(node => {
			const span = node.span;
			return {...node, get span() { reads += 1; return span; }};
		})};
		expect(procedureHasUnstructuredFlow(source, watched)).toBe(true);
		expect(reads).toBeLessThan(20);
		reads = 0;
		expect(procedureHasUnstructuredFlow(source, watched)).toBe(true);
		expect(reads).toBe(0);
	});

	it('finds a nested witness and keeps ordinary nested control flow structured', () => {
		for (const [inner, expected] of [['On Error Resume Next', true], ['x = 1', false]] as const) {
			const source = 'Sub P()\nIf x Then\nDo While y\n' + inner + '\nLoop\nEnd If\nEnd Sub';
			expect(procedureHasUnstructuredFlow(source, procedure(source).proc)).toBe(expected);
		}
	});

	it('changes activity on a reused procedure and retains a later active witness', () => {
		const source = 'Sub P()\n#If FLAG Then\nOn Error Resume Next\n#End If\nx = 1\nEnd Sub';
		const {module, proc} = procedure(source);
		for (const flag of [true, false, true]) {
			const activity = createConditionalActivityTracker(module, {compilerConstants: {FLAG: flag}});
			expect(procedureHasUnstructuredFlow(source, proc, activity)).toBe(flag);
		}
		const later = source.replace('x = 1', 'Resume Next');
		const parsed = procedure(later);
		const activity = createConditionalActivityTracker(parsed.module, {compilerConstants: {FLAG: false}});
		expect(procedureHasUnstructuredFlow(later, parsed.proc, activity)).toBe(true);
	});

	it('retains source-sensitive cache invalidation', () => {
		const source = 'Sub P()\nGoTo Done\nEnd Sub';
		const plain = source.replace('GoTo Done', 'x = 12345');
		const {proc} = procedure(source);
		expect(procedureHasUnstructuredFlow(source, proc)).toBe(true);
		expect(procedureHasUnstructuredFlow(plain, proc)).toBe(false);
		expect(procedureHasUnstructuredFlow(source, proc)).toBe(true);
	});

	it('shares label traversal while preserving inactive filtering and source order', () => {
		const source = 'Sub P()\n#If FLAG Then\nFirst:\nIf x Then\nGoTo First\nEnd If\n#Else\nOther:\nResume Other\n#End If\nEnd Sub';
		const {module, proc} = procedure(source);
		for (const flag of [true, false]) {
			const activity = createConditionalActivityTracker(module, {compilerConstants: {FLAG: flag}});
			const expected = flag ? 'First' : 'Other';
			expect(collectProcedureLabelDeclarations(source, proc, activity).map(label => label.text)).toEqual([expected]);
			expect(collectProcedureLabelReferences(source, proc, activity).map(label => label.text)).toEqual([expected]);
		}
	});
});
