// XML documentation model (developer-defined IntelliSense metadata).
//
// XLIDE lets a developer annotate their own procedures, types, and module
// members with Visual-Studio-style XML documentation comments (lines beginning
// with `'''`), and lets a developer ship the same vocabulary in external
// metadata files so symbols the curated library does not know about (or that
// the developer wants to re-describe) gain rich tooltips and call tips.
//
// This module is the host-agnostic data model shared by the inline parser
// (docComment.ts), the external-file parser (externalDoc.ts), and the lookup
// registry (docRegistry.ts). It carries no `vscode` dependency. See
// docs/vba-doc-comments.md for the full standard and usage paths.

/** Where a {@link VbaDoc} was authored. */
export type VbaDocSource = 'inline' | 'external';

/** Documentation for a single named parameter (`<param name="x">...`). */
export interface VbaDocParam {
	/** Declared parameter name, matched case-insensitively to the signature. */
	name: string;
	/** Rendered description text (entities decoded, whitespace collapsed). */
	text: string;
	/** Optional type hint from `type="..."` documentation metadata. */
	type?: string;
	/** Optional unit hint from `unit="..."` documentation metadata. */
	unit?: string;
	/** Optional value/category hint from `value="..."` documentation metadata. */
	value?: string;
}

/**
 * Parsed XML documentation for one symbol. Every field is optional except the
 * provenance marker, so a doc block may carry only a summary, only parameter
 * notes, or any combination of the supported tags.
 */
export interface VbaDoc {
	/** `<summary>` - the one-line description shown first in tooltips. */
	summary?: string;
	/** `<param name="...">` entries, in document order. */
	params: VbaDocParam[];
	/** `<returns>` - description of a function's return value. */
	returns?: string;
	/** Optional return type hint from `<returns type="...">`. */
	returnsType?: string;
	/** Optional return unit hint from `<returns unit="...">`. */
	returnsUnit?: string;
	/** Optional return value/category hint from `<returns value="...">`. */
	returnsValue?: string;
	/** `<remarks>` - extended notes shown below the summary. */
	remarks?: string;
	/** `<example>` - a usage example, rendered as a VBA code block. */
	example?: string;
	/**
	 * `<signature>` - an explicit call-tip signature. Used by external metadata
	 * to give a signature to a symbol the curated library cannot resolve. Ignored
	 * when a real signature is available from source or the host model.
	 */
	signature?: string;
	/** Whether this doc came from an inline `'''` comment or an external file. */
	source: VbaDocSource;
}

/** True when a doc block carries any human-readable prose worth showing. */
export function hasDocContent(doc: VbaDoc | undefined): doc is VbaDoc {
	return (
		!!doc &&
		(!!doc.summary ||
			!!doc.returns ||
			!!doc.returnsType ||
			!!doc.returnsUnit ||
			!!doc.returnsValue ||
			!!doc.remarks ||
			!!doc.example ||
			doc.params.length > 0)
	);
}

/**
 * Renders a doc block to Markdown for a hover tooltip: summary, then parameter
 * notes, returns, remarks, and an example code block. Parameter notes are
 * included here because hovers have no per-parameter surface.
 */
export function renderDocMarkdown(doc: VbaDoc): string {
	const parts: string[] = [];
	if (doc.summary) {
		parts.push(doc.summary);
	}
	if (doc.params.length > 0) {
		const lines = doc.params.map((p) => `- \`${p.name}\`${renderHintSuffix(p)}: ${p.text}`);
		parts.push(['**Parameters:**', ...lines].join('  \n'));
	}
	if (doc.returns) {
		parts.push(`**Returns${renderReturnHintSuffix(doc)}:** ${doc.returns}`);
	}
	if (doc.remarks) {
		parts.push(`**Remarks:** ${doc.remarks}`);
	}
	if (doc.example) {
		parts.push(`**Example:**\n\n\`\`\`vba\n${doc.example}\n\`\`\``);
	}
	return parts.join('\n\n');
}

/** Renders one parameter's doc for signature-help parameter details. */
export function renderParamDocMarkdown(param: VbaDocParam): string {
	const hints = renderHintList(param);
	if (hints.length === 0) {
		return param.text;
	}
	return `${hints.join(', ')}\n\n${param.text}`;
}

function renderHintSuffix(param: Pick<VbaDocParam, 'type' | 'unit' | 'value'>): string {
	const hints = renderHintList(param);
	return hints.length > 0 ? ` (${hints.join(', ')})` : '';
}

function renderReturnHintSuffix(doc: VbaDoc): string {
	const hints = renderHintList({
		type: doc.returnsType,
		unit: doc.returnsUnit,
		value: doc.returnsValue,
	});
	return hints.length > 0 ? ` (${hints.join(', ')})` : '';
}

function renderHintList(param: Pick<VbaDocParam, 'type' | 'unit' | 'value'>): string[] {
	const hints: string[] = [];
	if (param.type) {
		hints.push(`As ${param.type}`);
	}
	if (param.unit) {
		hints.push(`unit: ${param.unit}`);
	}
	if (param.value) {
		hints.push(`value: ${param.value}`);
	}
	return hints;
}

/**
 * Renders the summary-level Markdown for a signature-help call tip: summary,
 * returns, and remarks. Parameter notes are intentionally omitted here because
 * signature help shows them per-parameter.
 */
export function renderSignatureDocMarkdown(doc: VbaDoc): string {
	const parts: string[] = [];
	if (doc.summary) {
		parts.push(doc.summary);
	}
	if (doc.returns) {
		parts.push(`**Returns:** ${doc.returns}`);
	}
	if (doc.remarks) {
		parts.push(`**Remarks:** ${doc.remarks}`);
	}
	return parts.join('\n\n');
}
