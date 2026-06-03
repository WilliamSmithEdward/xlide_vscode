import { describe, expect, it } from 'vitest';
import {
	collectProcedureLabelReferences,
	collectProcedureLabels,
	parseModule,
	resolveProcedureLabelCompletions,
	resolveProcedureLabelDefinitionAt,
	type ProcedureNode,
} from '../src/analyzer';

function withCaret(marked: string): { source: string; offset: number } {
	const offset = marked.indexOf('|');
	if (offset < 0) {
		throw new Error('Missing caret marker.');
	}
	return { source: marked.slice(0, offset) + marked.slice(offset + 1), offset };
}

function onlyProcedure(source: string): ProcedureNode {
	const proc = parseModule(source).members.find((member): member is ProcedureNode =>
		member.kind === 'Procedure'
	);
	if (!proc) {
		throw new Error('Expected a procedure.');
	}
	return proc;
}

function completionLabels(marked: string): string[] {
	const { source, offset } = withCaret(marked);
	return resolveProcedureLabelCompletions(source, offset).map((item) => item.label);
}

describe('procedure label surface', () => {
	it('collects procedure-local labels and control-flow references', () => {
		const source = [
			'Sub T()',
			'    GoTo Done',
			'Start:',
			'    GoSub 20',
			'20:',
			'Done:',
			'End Sub',
		].join('\n');
		const proc = onlyProcedure(source);

		expect([...collectProcedureLabels(source, proc).values()].map((label) => label.text))
			.toEqual(['Start', '20', 'Done']);
		expect(collectProcedureLabelReferences(source, proc).map((ref) => ref.text))
			.toEqual(['Done', '20']);
	});

	it('offers labels at direct branch targets', () => {
		const labels = completionLabels([
			'Sub T()',
			'    GoTo |',
			'Start:',
			'20:',
			'End Sub',
		].join('\n'));

		expect(labels).toEqual(['Start', '20']);
	});

	it('offers labels in On Error and On n branch targets', () => {
		expect(completionLabels([
			'Sub T(ByVal n As Long)',
			'    On Error GoTo H|',
			'Handler:',
			'End Sub',
		].join('\n'))).toEqual(['Handler']);

		expect(completionLabels([
			'Sub T(ByVal n As Long)',
			'    On n GoTo First, |',
			'First:',
			'Second:',
			'End Sub',
		].join('\n'))).toEqual(['First', 'Second']);
	});

	it('does not treat valid On Error non-label forms as label targets', () => {
		expect(completionLabels('Sub T()\n    On Error GoTo 0|\nHandler:\nEnd Sub\n')).toEqual([]);
		expect(completionLabels('Sub T()\n    On Error GoTo -1|\nHandler:\nEnd Sub\n')).toEqual([]);
		expect(completionLabels('Sub T()\n    On Error Resume Next|\nHandler:\nEnd Sub\n')).toEqual([]);
	});

	it('resolves label references to same-procedure declarations', () => {
		const { source, offset } = withCaret([
			'Sub T()',
			'    GoTo Do|ne',
			'Done:',
			'End Sub',
		].join('\n'));

		const resolved = resolveProcedureLabelDefinitionAt(source, offset);
		expect(resolved?.label.text).toBe('Done');
		expect(source.slice(resolved!.label.span.start, resolved!.label.span.end)).toBe('Done');
		expect(resolved?.procedure.name).toBe('T');
	});

	it('resolves numeric line label references', () => {
		const { source, offset } = withCaret([
			'Sub T()',
			'    GoTo 1|0',
			'10:',
			'End Sub',
		].join('\n'));

		const resolved = resolveProcedureLabelDefinitionAt(source, offset);
		expect(resolved?.label.text).toBe('10');
		expect(source.slice(resolved!.label.span.start, resolved!.label.span.end)).toBe('10');
	});
});
