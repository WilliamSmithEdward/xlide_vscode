import type { ModuleNode } from './nodes';

/** Try one compatible snapshot; failed probes must not scan every historical class. */
export function incrementalModuleParseFromCache(
	source: string,
	snapshots: readonly { source: string; module: ModuleNode; hasDirectives?: boolean }[],
	parseFresh: (source: string) => ModuleNode,
): { module: ModuleNode; snapshot: { hasDirectives?: boolean } } | undefined {
	const previous = snapshots.find(entry => source.length >= 16_384 &&
		Math.abs(source.length - entry.source.length) <= 256 && source.slice(0, 128) === entry.source.slice(0, 128));
	const module = previous && incrementalModuleParse(source, previous.source, previous.module, parseFresh);
	return module && previous ? { module, snapshot: previous } : undefined;
}

/**
 * Immutable AST rebasing. Parser trees contain plain objects, arrays and spans,
 * and share nodes: an If block lists each statement in its branch and again in
 * its flat body. So each node is copied once and stays shared; copying it again
 * at every reference doubled the work with each level of nested If blocks.
 */
function rebase<T>(value: T, delta: number, copies: Map<object, unknown>): T {
	if (delta === 0 || value === null || typeof value !== 'object') { return value; }
	const done = copies.get(value);
	if (done !== undefined) { return done as T; }
	if (Array.isArray(value)) {
		const copy = new Array<unknown>(value.length);
		copies.set(value, copy);
		for (let i = 0; i < value.length; i++) { copy[i] = rebase(value[i], delta, copies); }
		return copy as T;
	}
	const record = value as Record<string, unknown>;
	if (typeof record.start === 'number' && typeof record.end === 'number') {
		const span = { ...record, start: record.start + delta, end: record.end + delta };
		copies.set(value, span);
		return span as T;
	}
	const copy: Record<string, unknown> = {};
	copies.set(value, copy);
	for (const key of Object.keys(record)) { copy[key] = rebase(record[key], delta, copies); }
	return copy as T;
}

/** Reparse a small physical-line edit inside an otherwise unchanged closed procedure. */
export function incrementalModuleParse(
	source: string,
	previousSource: string,
	previous: ModuleNode,
	parseFresh: (source: string) => ModuleNode,
): ModuleNode | undefined {
	// Small modules are already cheap. Different files, large replacements,
	// headers and newlines retain the full parser's recovery behavior.
	if (source.length < 16_384 || Math.abs(source.length - previousSource.length) > 256 ||
		source.slice(0, 128) !== previousSource.slice(0, 128)) { return undefined; }
	let start = 128;
	const commonEnd = Math.min(source.length, previousSource.length);
	while (start < commonEnd && source.charCodeAt(start) === previousSource.charCodeAt(start)) { start++; }
	let oldEnd = previousSource.length;
	let newEnd = source.length;
	while (oldEnd > start && newEnd > start &&
		previousSource.charCodeAt(oldEnd - 1) === source.charCodeAt(newEnd - 1)) { oldEnd--; newEnd--; }
	if (oldEnd - start > 256 || newEnd - start > 256 ||
		/[\r\n]/.test(previousSource.slice(start, oldEnd) + source.slice(start, newEnd))) { return undefined; }
	const index = previous.members.findIndex(member => member.kind === 'Procedure' && member.closed &&
		member.span.start < start && member.span.end > oldEnd);
	if (index < 0) { return undefined; }
	const original = previous.members[index];
	const text = previousSource.slice(original.span.start, original.span.end);
	const firstLineEnd = text.search(/[\r\n]/);
	const lastLineStart = Math.max(text.lastIndexOf('\n'), text.lastIndexOf('\r')) + 1;
	if (firstLineEnd < 0 || start <= original.span.start + firstLineEnd ||
		oldEnd >= original.span.start + lastLineStart || /(?:^|:)[ \t]*#/m.test(text)) { return undefined; }
	// A pre-existing structural error can make a procedure depend on recovery
	// state outside its span. Only independently parseable bodies qualify.
	if (previous.diagnostics.some(item => item.span.start < original.span.end && item.span.end >= original.span.start)) {
		return undefined;
	}
	const delta = source.length - previousSource.length;
	const replacementText = source.slice(original.span.start, original.span.end + delta);
	if (/(?:^|:)[ \t]*#/m.test(replacementText)) { return undefined; }
	const parsed = parseFresh(replacementText);
	const replacement = parsed.members[0];
	if (parsed.members.length !== 1 || replacement?.kind !== 'Procedure' || !replacement.closed ||
		replacement.span.start !== 0 || replacement.span.end !== replacementText.length) { return undefined; }
	// One copy map per shift: the replacement moves to the procedure's start, the tail by the edit's length.
	const replacementCopies = new Map<object, unknown>();
	const tailCopies = new Map<object, unknown>();
	return {
		...previous,
		span: { start: 0, end: source.length },
		members: previous.members.map((member, i) => i < index ? member :
			i === index ? rebase(replacement, original.span.start, replacementCopies) : rebase(member, delta, tailCopies)),
		diagnostics: [
			...previous.diagnostics.filter(item => item.span.end < original.span.start),
			...parsed.diagnostics.map(item => rebase(item, original.span.start, replacementCopies)),
			...previous.diagnostics.filter(item => item.span.start >= original.span.end).map(item => rebase(item, delta, tailCopies)),
		],
	};
}
