// Inline XML documentation-comment parser.
//
// A documentation comment is a run of contiguous lines whose first non-space
// characters are `'''` (three apostrophes), sitting directly above a procedure,
// type, enum, Declare, or module-level variable declaration - mirroring the
// Visual Studio C# `///` convention. The body is a fragment of XML using the
// tag vocabulary in docModel.ts (<summary>, <param>, <returns>, <remarks>,
// <example>). Parsing is intentionally lenient (regex-based, no XML dependency)
// so a partially written or slightly malformed block still yields useful text.
//
// Pure analyzer code: no `vscode` dependency. See docs/vba-doc-comments.md.

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
	const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'i');
	const m = re.exec(body);
	return m ? m[1] : undefined;
}

/** Extracts every `<param name="...">...</param>` entry, in document order. */
function extractParams(body: string): VbaDocParam[] {
	const out: VbaDocParam[] = [];
	const re = /<param\s+name\s*=\s*"([^"]*)"\s*>([\s\S]*?)<\/param>/gi;
	let m: RegExpExecArray | null;
	while ((m = re.exec(body)) !== null) {
		const name = decodeEntities(m[1]).trim();
		if (name) {
			out.push({ name, text: collapse(m[2]) });
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
	const returns = firstTag(body, 'returns');
	if (returns !== undefined) {
		doc.returns = collapse(returns);
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
		let rest = trimmed.slice(3);
		if (rest.startsWith(' ')) {
			rest = rest.slice(1);
		}
		docLines.push(rest);
	}
	if (docLines.length === 0) {
		return undefined;
	}
	docLines.reverse();
	const doc = parseDocBody(docLines.join('\n'), 'inline');
	return doc.summary ||
		doc.params.length > 0 ||
		doc.returns ||
		doc.remarks ||
		doc.example ||
		doc.signature
		? doc
		: undefined;
}
