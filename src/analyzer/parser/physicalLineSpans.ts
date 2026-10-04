import type { Span } from './nodes';

export function firstLineBreakAtOrAfter(source: string, start: number): number {
	for (let i = start; i < source.length; i++) {
		const ch = source[i];
		if (ch === '\n' || ch === '\r') {
			return i;
		}
	}
	return -1;
}

/** The block's header line, with the lines a ` _` continues it onto. */
export function blockHeaderLineSpan(source: string, span: Span): Span {
	let nl = firstLineBreakAtOrAfter(source, span.start);
	while (nl >= 0 && nl <= span.end && endsInContinuation(source, span.start, nl)) {
		const next = source[nl] === '\r' && source[nl + 1] === '\n' ? nl + 2 : nl + 1;
		nl = firstLineBreakAtOrAfter(source, next);
	}
	if (nl < 0 || nl > span.end) {
		return span;
	}
	return { start: span.start, end: nl };
}

function endsInContinuation(source: string, start: number, nl: number): boolean {
	let i = nl - 1;
	while (i >= start && (source[i] === ' ' || source[i] === '\t')) {
		i--;
	}
	return i > start && source[i] === '_' && (source[i - 1] === ' ' || source[i - 1] === '\t');
}
