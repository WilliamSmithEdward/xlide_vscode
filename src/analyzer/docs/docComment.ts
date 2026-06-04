// Inline XML documentation-comment parser.
//
// A documentation comment is a run of contiguous lines whose first non-space
// characters are `'''` (three apostrophes), sitting directly above a procedure,
// type, enum, Declare, or module-level variable declaration - mirroring the
// Visual Studio C# `///` convention. Object-module docs use the same block above
// the first Option directive because VBA has no source-level class declaration.
// The body is a fragment of XML using the tag vocabulary in docModel.ts
// (<summary>, <param>, <returns>, <remarks>, <example>). Parsing is
// intentionally lenient (regex-based, no XML dependency) so a partially written
// or slightly malformed block still yields useful text.
//
// Pure analyzer code: no `vscode` dependency. See user_guides/vba-doc-comments.md.

import { VbaDoc, VbaDocParam, VbaDocSource } from './docModel';

/** Decodes the five predefined XML entities. */
function decodeEntities(text: string): string {
	return text
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&');
}

/** Trims and collapses internal whitespace runs (including newlines) to a space. */
function collapse(text: string): string {
	return decodeEntities(text).replace(/\s+/g, ' ').trim();
}

/** Trims surrounding blank lines but preserves internal layout (for examples). */
function dedent(text: string): string {
	const decoded = decodeEntities(text).replace(/\r\n/g, '\n');
	return decoded.replace(/^\n+/, '').replace(/\s+$/, '');
}

/** Extracts the inner text of the first `<tag>...</tag>` in `body`, if present. */
function firstTag(body: string, tag: string): string | undefined {
	const m = firstTagMatch(body, tag);
	return m?.body;
}

function firstTagMatch(body: string, tag: string): { attrs: string; body: string } | undefined {
	const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'i');
	const m = re.exec(body);
	if (!m) {
		return undefined;
	}
	const open = new RegExp(`<${tag}([^>]*)>`, 'i').exec(m[0]);
	return { attrs: open?.[1] ?? '', body: m[1] };
}

function attrsOf(raw: string): Map<string, string> {
	const out = new Map<string, string>();
	const re = /([A-Za-z_][A-Za-z0-9_-]*)\s*=\s*"([^"]*)"/g;
	let m: RegExpExecArray | null;
	while ((m = re.exec(raw)) !== null) {
		out.set(m[1].toLowerCase(), decodeEntities(m[2]).trim());
	}
	return out;
}

/** Extracts every `<param name="...">...</param>` entry, in document order. */
function extractParams(body: string): VbaDocParam[] {
	const out: VbaDocParam[] = [];
	const re = /<param\b([^>]*)>([\s\S]*?)<\/param>/gi;
	let m: RegExpExecArray | null;
	while ((m = re.exec(body)) !== null) {
		const attrs = attrsOf(m[1]);
		const name = attrs.get('name') ?? '';
		if (name) {
			const param: VbaDocParam = { name, text: collapse(m[2]) };
			const type = attrs.get('type');
			const unit = attrs.get('unit');
			const value = attrs.get('value');
			if (type) {
				param.type = type;
			}
			if (unit) {
				param.unit = unit;
			}
			if (value) {
				param.value = value;
			}
			out.push(param);
		}
	}
	return out;
}

const HAS_TAG_RE = /<(summary|param|returns|remarks|example|signature)\b/i;

/**
 * Parses an XML documentation body (the inner text shared by inline comments and
 * external `<member>` entries) into a {@link VbaDoc}. When the body contains no
 * recognised tags, the whole trimmed text becomes the summary so a plain-text
 * note still produces a tooltip.
 */
export function parseDocBody(body: string, source: VbaDocSource): VbaDoc {
	const doc: VbaDoc = { params: [], source };
	if (!HAS_TAG_RE.test(body)) {
		const plain = collapse(body);
		if (plain) {
			doc.summary = plain;
		}
		return doc;
	}
	const summary = firstTag(body, 'summary');
	if (summary !== undefined) {
		doc.summary = collapse(summary);
	}
	doc.params = extractParams(body);
	const returns = firstTagMatch(body, 'returns');
	if (returns !== undefined) {
		doc.returns = collapse(returns.body);
		const attrs = attrsOf(returns.attrs);
		const type = attrs.get('type');
		const unit = attrs.get('unit');
		const value = attrs.get('value');
		if (type) {
			doc.returnsType = type;
		}
		if (unit) {
			doc.returnsUnit = unit;
		}
		if (value) {
			doc.returnsValue = value;
		}
	}
	const remarks = firstTag(body, 'remarks');
	if (remarks !== undefined) {
		doc.remarks = collapse(remarks);
	}
	const example = firstTag(body, 'example');
	if (example !== undefined) {
		doc.example = dedent(example);
	}
	const signature = firstTag(body, 'signature');
	if (signature !== undefined) {
		doc.signature = collapse(signature);
	}
	return doc;
}

function hasAnyDocContent(doc: VbaDoc): boolean {
	return !!doc.summary ||
		doc.params.length > 0 ||
		!!doc.returns ||
		!!doc.remarks ||
		!!doc.example ||
		!!doc.signature;
}

function docFromLines(docLines: readonly string[]): VbaDoc | undefined {
	if (docLines.length === 0) {
		return undefined;
	}
	const doc = parseDocBody(docLines.join('\n'), 'inline');
	return hasAnyDocContent(doc) ? doc : undefined;
}

function stripDocPrefix(trimmed: string): string {
	let rest = trimmed.slice(3);
	if (rest.startsWith(' ')) {
		rest = rest.slice(1);
	}
	return rest;
}

interface SourceLine {
	text: string;
	start: number;
	end: number;
}

function sourceLines(source: string): SourceLine[] {
	const rawLines = source.split('\n');
	const out: SourceLine[] = [];
	let offset = 0;
	for (const raw of rawLines) {
		const text = raw.replace(/\r$/, '');
		out.push({
			text,
			start: offset,
			end: offset + raw.length,
		});
		offset += raw.length + 1;
	}
	return out;
}

function firstLineIndexAtOrAfter(lines: readonly SourceLine[], offset: number): number {
	const safeOffset = Math.max(0, offset);
	for (let i = 0; i < lines.length; i += 1) {
		const line = lines[i];
		if (line.start >= safeOffset) {
			return i;
		}
		if (safeOffset <= line.end) {
			return i + 1;
		}
	}
	return lines.length;
}

function isOrdinaryComment(trimmed: string): boolean {
	return trimmed.startsWith("'") && !trimmed.startsWith("'''");
}

function isModuleHeaderBoundary(trimmed: string): boolean {
	return trimmed === '' ||
		isOrdinaryComment(trimmed) ||
		/^Option\b/i.test(trimmed);
}

/**
 * Extracts a module-header documentation block from the top of a module. This
 * supports both object modules with a block directly above `Option Explicit`
 * and standard modules that have no `Option` directive by requiring the header
 * block to be visually separated from the first declaration by a blank line or
 * ordinary comment. A block immediately above a declaration remains declaration
 * documentation.
 */
export function extractModuleHeaderDoc(
	source: string,
	startOffset = 0,
): VbaDoc | undefined {
	const lines = sourceLines(source);
	let i = firstLineIndexAtOrAfter(lines, startOffset);
	while (i < lines.length) {
		const trimmed = lines[i].text.trimStart();
		if (trimmed === '' || isOrdinaryComment(trimmed)) {
			i += 1;
			continue;
		}
		break;
	}
	if (i >= lines.length || !lines[i].text.trimStart().startsWith("'''")) {
		return undefined;
	}

	const docLines: string[] = [];
	while (i < lines.length) {
		const trimmed = lines[i].text.trimStart();
		if (!trimmed.startsWith("'''")) {
			break;
		}
		docLines.push(stripDocPrefix(trimmed));
		i += 1;
	}

	if (i < lines.length) {
		const next = lines[i].text.trimStart();
		if (!isModuleHeaderBoundary(next)) {
			return undefined;
		}
	}

	return docFromLines(docLines);
}

/**
 * Scans upward from the start of a declaration to collect a contiguous run of
 * `'''` documentation-comment lines, and parses them into a {@link VbaDoc}.
 * Returns undefined when no such comment immediately precedes the declaration.
 *
 * @param source Full module source text.
 * @param declStart Offset of the first character of the declaration (its full
 *   span start, e.g. the `Public`/`Sub` keyword).
 */
export function extractLeadingDoc(
	source: string,
	declStart: number,
): VbaDoc | undefined {
	const before = source.slice(0, Math.max(0, declStart));
	const physical = before.split('\n');
	// Drop the partial current line (indentation before the declaration keyword).
	physical.pop();
	const docLines: string[] = [];
	for (let i = physical.length - 1; i >= 0; i -= 1) {
		const line = physical[i].replace(/\r$/, '');
		const trimmed = line.trimStart();
		if (!trimmed.startsWith("'''")) {
			break;
		}
		docLines.push(stripDocPrefix(trimmed));
	}
	docLines.reverse();
	return docFromLines(docLines);
}
