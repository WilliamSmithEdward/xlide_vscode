// XML documentation model (developer-defined IntelliSense metadata).
//
// XLIDE lets a developer annotate their own procedures, types, and module
// members with Visual-Studio-style XML documentation comments (lines beginning
// with `'''`), and lets a team ship the same vocabulary in external metadata
// files so symbols the curated library does not know about (or that the team
// wants to re-describe) gain rich tooltips and call tips.
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
		const lines = doc.params.map((p) => `- \`${p.name}\`: ${p.text}`);
		parts.push(['**Parameters:**', ...lines].join('  \n'));
	}
	if (doc.returns) {
		parts.push(`**Returns:** ${doc.returns}`);
	}
	if (doc.remarks) {
		parts.push(`**Remarks:** ${doc.remarks}`);
	}
	if (doc.example) {
		parts.push(`**Example:**\n\n\`\`\`vba\n${doc.example}\n\`\`\``);
	}
	return parts.join('\n\n');
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
