import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { PushFn } from '../src/analyzer/diagnostics/analysisContext';
import { checkDocComments } from '../src/analyzer/diagnostics/rules/docComments';
import { parseModule } from '../src/analyzer/parser/parseModule';

const counters = vi.hoisted(() => ({startReads: 0}));
vi.mock('../src/analyzer/docs/docComment', async (importOriginal) => {
	const actual = await importOriginal<typeof import('../src/analyzer/docs/docComment')>();
	return {...actual, leadingDocLines: (...args: Parameters<typeof actual.leadingDocLines>) =>
		actual.leadingDocLines(...args).map(line => {
			const start = line.start;
			return {...line, get start() { counters.startReads += 1; return start; }};
		})};
});

function findings(source: string): Parameters<PushFn>[] {
	const out: Parameters<PushFn>[] = [];
	checkDocComments(source, parseModule(source), undefined, (...args) => { out.push(args); });
	return out;
}

beforeEach(() => { counters.startReads = 0; });

describe('documentation quick-fix line lookup', () => {
	it('bounds line-start reads when creating many duplicate removal fixes', () => {
		const count = 200;
		const row = "''' <summary/>\n";
		const source = row.repeat(count) + 'Sub P()\nEnd Sub';
		const out = findings(source);
		expect(out).toHaveLength(count - 1);
		expect(counters.startReads).toBeLessThan(count * 50);
		out.forEach((item, i) => {
			expect(item[0]).toBe('docTagDuplicate');
			expect(item[3]!.docCommentFixes![0].edits).toEqual([
				{span: {start: (i + 1) * row.length, end: (i + 2) * row.length}, newText: ''},
			]);
		});
	});

	it.each(['\n', '\r\n'])('removes an attached next-line directive with its repeated tag (%j)', (eol) => {
		const directive = "' @xlide-analysis-disable-next-line DOC000";
		const source = ["''' <summary/>", directive, "''' <summary/>", "''' <remarks>r</remarks>", 'Sub P()', 'End Sub'].join(eol);
		const out = findings(source);
		expect(out).toHaveLength(1);
		expect(out[0][3]!.docCommentFixes![0].edits).toEqual([
			{span: {start: source.indexOf(directive), end: source.indexOf("''' <remarks>")}, newText: ''},
		]);
	});

	it('removes only a repeated tag when another tag shares its line', () => {
		const repeated = '<param name="A">second</param>';
		const source = ["''' <param name=\"A\">first</param>", "''' " + repeated + '<remarks>r</remarks>', 'Sub P(ByVal A As Long)', 'End Sub'].join('\n');
		const duplicate = findings(source).find(item => item[0] === 'docTagDuplicate')!;
		const start = source.indexOf(repeated);
		expect(duplicate[3]!.docCommentFixes![0].edits).toEqual([{span: {start, end: start + repeated.length}, newText: ''}]);
	});
});
