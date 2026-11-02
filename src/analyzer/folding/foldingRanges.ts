// Fold ranges for a module: the blocks the parser reads, each #If to its #End If,
// and the arms inside them, whatever the code's indentation (issue #1576). VS Code
// drops its own indentation folding once a provider answers, so the folds indentation
// gives are kept wherever no block claims the line or crosses them: a run of continued
// lines, a VB6 designer header, or a block still being typed.
//
// A block folds from its header to its closer, as End Sub does. An arm folds up to the
// line before the next arm or the closer, which stay visible. A block that is not
// closed has no range of its own.

import { parseModule } from '../parser/parseModule';
import type { BodyNode, ConditionalDirectiveNode, ModuleNode, Span } from '../parser/nodes';

/** Zero-based first and last line of a fold. */
export interface VbaFoldingRange {
	start: number;
	end: number;
}

/** Which range stays where two cross: a procedure or block over #If arms, which can open or close part of one. */
const Rank = { Indentation: 0, Directive: 1, Block: 2 } as const;
type Rank = (typeof Rank)[keyof typeof Rank];

interface Candidate extends VbaFoldingRange {
	rank: Rank;
	dropped?: boolean;
}

/** Spaces a tab advances to, for reading indentation: VS Code's default tab size. */
const TAB_SIZE = 4;
const CASE_ARM = /^Case\b/i;

/**
 * The fold ranges of a module, nested and sorted by first line.
 * `analysisSource` is what the parser reads, with the same lines as `text`.
 */
export function vbaFoldingRanges(analysisSource: string, text: string = analysisSource): VbaFoldingRange[] {
	const lineStarts = lineStartOffsets(analysisSource);
	const candidates = blockRanges(parseModule(analysisSource), analysisSource, lineStarts);
	candidates.push(...indentationRanges(text));
	return nestedRanges(candidates);
}

function lineStartOffsets(source: string): number[] {
	const starts = [0];
	for (let i = 0; i < source.length; i++) {
		const code = source.charCodeAt(i);
		if (code === 10 || (code === 13 && source.charCodeAt(i + 1) !== 10)) {
			starts.push(i + 1);
		}
	}
	return starts;
}

function lineAt(lineStarts: readonly number[], offset: number): number {
	let lo = 0;
	let hi = lineStarts.length - 1;
	while (lo < hi) {
		const mid = (lo + hi + 1) >> 1;
		if (lineStarts[mid] <= offset) { lo = mid; } else { hi = mid - 1; }
	}
	return lo;
}

function blockRanges(module: ModuleNode, source: string, lineStarts: readonly number[]): Candidate[] {
	const ranges: Candidate[] = [];
	const firstLine = (span: Span): number => lineAt(lineStarts, span.start);
	const lastLine = (span: Span): number => {
		let end = Math.min(span.end, source.length) - 1;
		while (end > span.start && /\s/.test(source[end])) { end--; }
		return lineAt(lineStarts, end);
	};
	const add = (start: number, end: number, rank: Rank = Rank.Block): void => {
		if (end > start) { ranges.push({ start, end, rank }); }
	};
	// Each arm's first line, then the closer's line: an arm runs up to the line before the next.
	const addArms = (arms: readonly number[], closer: number): void => {
		arms.forEach((arm, i) => add(arm, (i + 1 < arms.length ? arms[i + 1] : closer) - 1));
	};
	const directives: ConditionalDirectiveNode[] = [];

	const walkBody = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			switch (node.kind) {
				case 'ConditionalDirective':
					directives.push(node);
					break;
				case 'IfBlock':
					if (node.closed) {
						add(firstLine(node.span), lastLine(node.span));
						addArms(node.branches.slice(1).map((arm) => firstLine(arm.headerSpan)), lastLine(node.span));
					}
					walkBody(node.body);
					break;
				case 'SelectBlock':
					if (node.closed) {
						add(firstLine(node.span), lastLine(node.span));
						addArms(
							node.body.flatMap((arm) => (arm.kind === 'Statement' && CASE_ARM.test(arm.raw) ? [firstLine(arm.span)] : [])),
							lastLine(node.span),
						);
					}
					walkBody(node.body);
					break;
				case 'ForBlock':
				case 'DoBlock':
				case 'WhileBlock':
				case 'WithBlock':
					if (node.closed) { add(firstLine(node.span), lastLine(node.span)); }
					walkBody(node.body);
					break;
				default:
					break;
			}
		}
	};

	for (const member of module.members) {
		switch (member.kind) {
			case 'ConditionalDirective':
				directives.push(member);
				break;
			case 'Procedure':
				if (member.closed) { add(firstLine(member.span), lastLine(member.span)); }
				walkBody(member.body);
				break;
			case 'Type':
			case 'Enum':
				if (member.closed) { add(firstLine(member.span), lastLine(member.span)); }
				directives.push(...(member.directives ?? []));
				break;
			default:
				break;
		}
	}

	addDirectiveRanges(directives, firstLine, (start, end) => add(start, end, Rank.Directive));
	return ranges;
}

/** Each #If to its #End If, and each #ElseIf and #Else up to the line before the next. */
function addDirectiveRanges(
	directives: ConditionalDirectiveNode[],
	firstLine: (span: Span) => number,
	add: (start: number, end: number) => void,
): void {
	const open: number[][] = [];
	directives.sort((a, b) => a.span.start - b.span.start);
	for (const directive of directives) {
		const line = firstLine(directive.span);
		switch (directive.directiveKind) {
			case 'If':
				open.push([line]);
				break;
			case 'ElseIf':
			case 'Else': {
				const arms = open.at(-1);
				if (arms) {
					if (arms.length > 1) { add(arms[arms.length - 1], line - 1); }
					arms.push(line);
				}
				break;
			}
			case 'EndIf': {
				const arms = open.pop();
				if (arms) {
					if (arms.length > 1) { add(arms[arms.length - 1], line - 1); }
					add(arms[0], line);
				}
				break;
			}
			default:
				break;
		}
	}
}

/** VS Code's indentation folds: a line with deeper lines under it, up to the next line no deeper than it. */
function indentationRanges(text: string): Candidate[] {
	const ranges: Candidate[] = [];
	const open: Array<{ line: number; indent: number; deeper: boolean }> = [];
	const lines = text.split(/\r\n|\r|\n/);
	lines.forEach((line, index) => {
		const indent = indentOf(line);
		if (indent === undefined) { return; }
		while (open.length > 0 && open[open.length - 1].indent >= indent) {
			const closed = open.pop()!;
			if (closed.deeper) { ranges.push({ start: closed.line, end: index - 1, rank: Rank.Indentation }); }
		}
		if (open.length > 0) { open[open.length - 1].deeper = true; }
		open.push({ line: index, indent, deeper: false });
	});
	for (const left of open) {
		if (left.deeper) { ranges.push({ start: left.line, end: lines.length - 1, rank: Rank.Indentation }); }
	}
	return ranges;
}

/** The width of a line's leading blanks, or undefined for a blank line. */
function indentOf(line: string): number | undefined {
	let width = 0;
	for (let i = 0; i < line.length; i++) {
		const code = line.charCodeAt(i);
		if (code === 32) {
			width++;
		} else if (code === 9) {
			width += TAB_SIZE - (width % TAB_SIZE);
		} else {
			return width;
		}
	}
	return undefined;
}

/** One range per first line, properly nested. Where two cross, the higher rank stays, then the first. */
function nestedRanges(candidates: Candidate[]): VbaFoldingRange[] {
	candidates.sort((a, b) => a.start - b.start || b.rank - a.rank || b.end - a.end);
	const stack: Candidate[] = [];
	for (const range of candidates) {
		while (stack.length > 0 && stack[stack.length - 1].end < range.start) { stack.pop(); }
		while (stack.length > 0 && !range.dropped && (range.end > stack[stack.length - 1].end || range.start === stack[stack.length - 1].start)) {
			const top = stack[stack.length - 1];
			if (range.rank > top.rank && range.start !== top.start) {
				top.dropped = true;
				stack.pop();
			} else {
				range.dropped = true;
			}
		}
		if (!range.dropped) { stack.push(range); }
	}
	return candidates.filter((range) => !range.dropped).map(({ start, end }) => ({ start, end }));
}
