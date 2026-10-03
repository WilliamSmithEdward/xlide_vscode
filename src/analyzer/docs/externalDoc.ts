// External XML metadata-file parser.
//
// A metadata file describes documentation for one or more symbols using the
// same tag vocabulary as inline `'''` comments, wrapped in `<member>` elements
// keyed by symbol name - mirroring the .NET XML documentation-file layout:
//
//   <xlideDoc>
//     <member name="Module1.ComputeTax">
//       <summary>Returns the tax owed for an amount.</summary>
//       <param name="Amount">The pre-tax amount, in dollars.</param>
//       <returns>The tax owed, in dollars.</returns>
//     </member>
//     <member name="MsgBox">
//       <summary>Team note: prefer the Notify helper over raw MsgBox.</summary>
//     </member>
//   </xlideDoc>
//
// A `name` is either qualified (`Module.Symbol`) or bare (`Symbol`). Parsing is
// lenient and regex-based so a single malformed member never discards the rest
// of the file. Pure analyzer code: no `vscode` dependency. See
// user_guides/vba-doc-comments.md.

import { VbaDoc } from './docModel';
import { parseDocBody } from './docComment';

/** One symbol's documentation parsed from an external metadata file. */
export interface ExternalDocEntry {
	/** Qualified (`Module.Symbol`) or bare (`Symbol`) name key. */
	name: string;
	/** The parsed documentation. */
	doc: VbaDoc;
}

/**
 * Parses the text of an external metadata file into a list of doc entries.
 * Members without a usable name are skipped. Never throws.
 */
export function parseMetadataFile(xml: string): ExternalDocEntry[] {
	const out: ExternalDocEntry[] = [];
	const opening = /<member\s+name\s*=\s*"([^"]*)"\s*>/gi;
	const closing = /<\/member>/gi;
	let m: RegExpExecArray | null;
	while ((m = opening.exec(xml)) !== null) {
		const bodyStart = opening.lastIndex;
		closing.lastIndex = bodyStart;
		const close = closing.exec(xml);
		if (!close) {
			// No later opening can form a pair once no closing tag remains.
			break;
		}
		// Empty names still consume their whole pair, just as the old regex did.
		opening.lastIndex = closing.lastIndex;
		const name = m[1].trim();
		if (!name) {
			continue;
		}
		out.push({ name, doc: parseDocBody(xml.slice(bodyStart, close.index), 'external') });
	}
	return out;
}
