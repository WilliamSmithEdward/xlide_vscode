// The text shape of a VBA module, shared by every container that holds one:
// the hidden header (a class preamble or a form's designer block, then the
// `Attribute VB_*` lines), the visible body, and the procedure scan the
// explorer lists from. Workbook modules and VB6 module files are the same
// text here; only where the bytes live differs.

const ATTR_LINE_RE = /^Attribute\s+VB_/i;
// \p{L}, not \w: VBA identifiers may use any locale letter (a Russian
// project legitimately declares `Sub Proverka()` in Cyrillic), and the
// ASCII-only \w made such procedures vanish from the explorer tree. \p{M}
// after the first character because Thai and Devanagari build a letter from
// a base plus a combining mark, and stopping at the mark made those
// procedures vanish the same way.
const PROC_RE =
	/^[ \t]*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([\p{L}_][\p{L}\p{M}\p{N}_]*)\s*[(\r\n]/gimu;

export interface ProcedureEntry {
	name: string;
	kind: string;
	line: number;
}

/**
 * Split module source into (hidden header, visible body) exactly as the editor
 * surface expects: the VERSION/BEGIN/END class preamble plus the contiguous run
 * of `Attribute VB_*` lines are hidden; the body has leading blank lines removed.
 */
export function splitVbaSource(source: string): { header: string; body: string } {
	const lines = source.split(/(?<=\n)/); // keep line endings
	let i = 0;
	if (lines.length > 0 && /^VERSION\s+\d/i.test(lines[0])) {
		i++;
		// A class preamble opens with a bare BEGIN; a form's designer block
		// opens with `Begin {GUID} Name` and can NEST further Begin blocks for
		// its controls. Both are one balanced block, walked by depth - without
		// this, importing a .frm spliced the designer text into the code module.
		const opener = i < lines.length ? lines[i].replace(/[\r\n]+$/, '').trim() : '';
		if (/^BEGIN\b/i.test(opener)) {
			let depth = 1;
			i++;
			while (i < lines.length && depth > 0) {
				const line = lines[i].replace(/[\r\n]+$/, '').trim();
				if (/^Begin\b/i.test(line)) {
					depth++;
				} else if (/^End$/i.test(line)) {
					depth--;
				}
				i++;
			}
		}
	}
	while (i < lines.length && ATTR_LINE_RE.test(lines[i])) {
		i++;
	}
	return {
		header: lines.slice(0, i).join(''),
		body: lines.slice(i).join('').replace(/^[\r\n]+/, ''),
	};
}

/**
 * The end offset of the block a file opens with when it is a designer or
 * class preamble: `VERSION n` then `Begin`/`BEGIN` ... `End`/`END`, nested
 * once per control. `BeginProperty` ... `EndProperty` blocks inside it are
 * property groups (a Font), not nesting, and pass through. Zero when the
 * file opens with no such block, or the block never closes - in which case
 * nothing is treated as header, because hiding code is worse than showing a
 * header.
 */
export function designerHeaderEnd(text: string): number {
	const lines = text.split(/(?<=\n)/);
	let i = 0;
	while (i < lines.length && lines[i].trim() === '') {
		i++;
	}
	if (i >= lines.length || !/^\s*VERSION\s+\d/i.test(lines[i])) {
		return 0;
	}
	i++;
	while (i < lines.length && lines[i].trim() === '') {
		i++;
	}
	if (i >= lines.length || !/^\s*Begin\b/i.test(lines[i])) {
		return 0;
	}
	let depth = 0;
	let offset = lines.slice(0, i).reduce((sum, line) => sum + line.length, 0);
	for (; i < lines.length; i++) {
		const line = lines[i];
		offset += line.length;
		if (/^\s*Begin\b/i.test(line)) {
			depth++;
		} else if (/^\s*End\s*$/i.test(line)) {
			depth--;
			if (depth === 0) {
				return offset;
			}
		}
	}
	return 0;
}

/**
 * The text with its designer or class preamble turned to spaces, line
 * breaks kept: the parser sees VBA only, and every offset is still the
 * file's own.
 */
export function blankDesignerHeader(text: string): string {
	const end = designerHeaderEnd(text);
	if (end === 0) {
		return text;
	}
	return text.slice(0, end).replace(/[^\r\n]/g, ' ') + text.slice(end);
}

export function joinVbaSource(header: string, body: string): string {
	if (!header) { return body; }
	return `${header.replace(/[\r\n]+$/, '')}\r\n${body}`;
}

/**
 * The value of an `Attribute VB_*` header line.
 *
 * Both spellings count. The VBE quotes a string value (`VB_Name = "Ticket"`)
 * and leaves a boolean bare (`VB_PredeclaredId = True`), and reading only the
 * quoted form made every boolean attribute answer the empty string - which
 * silently disabled the document-module fallback in the classifier, whose
 * whole job is to recognise a host module by `PredeclaredId` and `Exposed`
 * both being True.
 */
export function attributeValue(source: string, attribute: string): string {
	const re = new RegExp(`^\\s*Attribute\\s+${attribute}\\s*=\\s*(?:"([^"]*)"|([^\\r\\n]*))`, 'im');
	const match = re.exec(source);
	return (match?.[1] ?? match?.[2] ?? '').trim();
}

/**
 * The procedures a module body declares, with 1-based lines. Line numbers
 * come from one forward pass: counting them per match with
 * `body.slice(0, m.index).split('\n')` re-walked the module from the start
 * every time - quadratic, and on a 26,000-line class it made expanding the
 * module in the explorer cost about 370 ms of which 1.5 ms was the search.
 * Matches arrive in increasing order, so each character is counted once.
 */
export function listProcedures(body: string): ProcedureEntry[] {
	const out: ProcedureEntry[] = [];
	PROC_RE.lastIndex = 0;
	let m: RegExpExecArray | null;
	let scanned = 0;
	let line = 1;
	while ((m = PROC_RE.exec(body)) !== null) {
		for (; scanned < m.index; scanned += 1) {
			if (body.charCodeAt(scanned) === 10) {
				line += 1;
			}
		}
		out.push({
			name: m[2],
			kind: m[1].replace(/\s+/g, ' ').trim(),
			line,
		});
	}
	return out;
}
