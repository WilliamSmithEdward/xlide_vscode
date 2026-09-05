/**
 * Which procedure a line belongs to, the way the VBE's
 * `CodeModule.ProcOfLine` answers it:
 *
 * - a procedure runs from its header to its End line, and also takes the
 *   comment and blank lines above its header, back to the previous
 *   procedure's End;
 * - the declarations section is everything above the first procedure's own
 *   leading comments, up to and including the last line there that is code;
 * - lines after the last procedure's End belong to that procedure.
 *
 * The last two rules mean no End line has to be found at all: each procedure
 * simply owns everything from its own leading run up to the next procedure's.
 *
 * Parity with xlide_vbide's procedureat
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/66).
 */

/**
 * Anchored per line, and deliberately not matching `Declare`: a
 * `Private Declare Function` is a declaration, not a procedure header, and
 * `Declare` is not one of the modifiers the pattern lets through.
 */
const PROCEDURE_HEADER_RE =
	/^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*(Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+([\p{L}_][\p{L}\p{M}\p{N}_]*)[ \t]*(?:\(|$)/iu;

const COMMENT_LINE_RE = /^[ \t]*(?:'|Rem(?=[ \t:]|$))/i;

/** CRLF, lone CR and lone LF. */
const LINE_BREAK_RE = /\r\n|\r|\n/;

export interface VbaProcedureRange {
	/** `Sub`, `Function`, `Property Get`, ... as the source spells it. */
	kind: string;
	name: string;
	/** 0-based, inclusive: the first line the procedure owns. */
	firstLine: number;
	/** 0-based, inclusive: the last line the procedure owns. */
	lastLine: number;
}

/**
 * Every procedure in a module's body, with the lines each one owns. Lines are
 * 0-based, matching a VS Code document.
 */
export function vbaProcedureRanges(source: string): VbaProcedureRange[] {
	const lines = source.split(LINE_BREAK_RE);
	const headers: Array<{ kind: string; name: string; line: number }> = [];
	for (let i = 0; i < lines.length; i += 1) {
		const match = PROCEDURE_HEADER_RE.exec(lines[i]);
		if (match) {
			headers.push({ kind: match[1].replace(/\s+/g, ' ').trim(), name: match[2], line: i });
		}
	}

	// The comment and blank lines above a header belong to it, but never past
	// the previous header - a stray line of code between two procedures stays
	// with the one above it. Walked back from each header rather than tracked
	// forward, so the (much more expensive) lead-in test runs on the handful of
	// lines above a procedure instead of on every line of the module.
	const firstLines = headers.map((header, index) => {
		const floor = index === 0 ? 0 : headers[index - 1].line + 1;
		let first = header.line;
		while (first > floor && isLeadIn(lines[first - 1])) {
			first -= 1;
		}
		return first;
	});

	return headers.map((header, index) => ({
		kind: header.kind,
		name: header.name,
		firstLine: firstLines[index],
		lastLine: index + 1 < headers.length ? firstLines[index + 1] - 1 : Math.max(lines.length - 1, 0),
	}));
}

/**
 * The procedure a 0-based line sits in, or undefined for the declarations
 * section - which is everything above the first procedure's leading run.
 */
export function vbaProcedureAtLine(
	ranges: readonly VbaProcedureRange[],
	line: number,
): VbaProcedureRange | undefined {
	return ranges.find((range) => line >= range.firstLine && line <= range.lastLine);
}

/** The VBE's own wording for a caret above the first procedure. */
export const VBA_DECLARATIONS_LABEL = '(Declarations)';

/** How the status bar names a procedure, or the lack of one. */
export function vbaProcedureLabel(range: VbaProcedureRange | undefined): string {
	return range ? `${range.kind} ${range.name}` : VBA_DECLARATIONS_LABEL;
}

/** What the status bar shows for a caret on this line of a module. */
export function vbaProcedureLabelAtLine(source: string, line: number): string {
	return vbaProcedureLabel(vbaProcedureAtLine(vbaProcedureRanges(source), line));
}

function isLeadIn(line: string): boolean {
	return line.trim() === '' || COMMENT_LINE_RE.test(line);
}
