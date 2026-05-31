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
// docs/vba-doc-comments.md.

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
	const re = /<member\s+name\s*=\s*"([^"]*)"\s*>([\s\S]*?)<\/member>/gi;
	let m: RegExpExecArray | null;
	while ((m = re.exec(xml)) !== null) {
		const name = m[1].trim();
		if (!name) {
			continue;
		}
		out.push({ name, doc: parseDocBody(m[2], 'external') });
	}
	return out;
}
