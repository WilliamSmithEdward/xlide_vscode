/**
 * The Rubberduck `@Folder` annotation: a comment in a module's declarations
 * section that says where the module belongs in a folder layout.
 *
 *     '@Folder("Accounts.Ledger")
 *     Option Explicit
 *
 * Read leniently, written one way. All of these mean the same folder:
 *
 *     '@Folder "Accounts.Ledger"
 *     '@Folder(Accounts.Ledger)
 *     '@Folder Accounts.Ledger
 *     '@folder("accounts.ledger")
 *
 * Parity with xlide_vbide's FolderAnnotation, which owns the same rules.
 */

/**
 * The name ends at whitespace, a bracket, a quote, or the line's end, so
 * `@Folders` and `@Folder-ish` are prose and match nothing.
 */
const FOLDER_TAG_RE = /@folder(?=[\s()"]|$)/i;

/**
 * A procedure header ends the declarations section. `Declare` is not a header:
 * it sits between the modifier and the `Function`, so the pattern's Sub /
 * Function / Property has to follow the modifiers directly.
 */
const PROCEDURE_HEADER_RE =
	/^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+/i;

/** CRLF, lone CR and lone LF, matching what `String.split` on the same pattern does. */
const LINE_BREAK_RE = /\r\n|\r|\n/g;

/** `Rem` opening a comment: the whole word, and nothing glued to it. */
const REM_COMMENT_RE = /^rem(?=[\s:]|$)/i;

export interface FolderAnnotationRead {
	/** The normalized dotted path. Absent when the module names no folder. */
	folder?: string;
	/**
	 * The read reached a verdict: it found an annotation, or it ran into the
	 * first procedure header, which is where the declarations section stops.
	 * False means the text ran out first, so a caller holding only a prefix of
	 * the module still has somewhere to look.
	 */
	complete: boolean;
}

export interface FolderAnnotationOptions {
	/**
	 * The text is a prefix of the module rather than the whole of it, so its
	 * last line may be cut mid-word. That line is skipped: reading
	 * `'@Folder("Accounts.Led` as a folder named `Accounts.Led` would be worse
	 * than reading nothing, because nothing is what asks for the longer read.
	 */
	truncated?: boolean;
}

/**
 * The folder a module's declarations section names, if any. The first
 * annotation wins; later ones are ignored.
 */
export function readFolderAnnotation(
	text: string,
	options: FolderAnnotationOptions = {},
): FolderAnnotationRead {
	// Walked rather than split: the annotation is a top-of-module convention, so
	// the read almost always stops within a few lines, and splitting first would
	// mean allocating every line of a large module to look at four of them.
	LINE_BREAK_RE.lastIndex = 0;
	let start = 0;
	let match: RegExpExecArray | null;
	for (;;) {
		match = LINE_BREAK_RE.exec(text);
		const line = text.slice(start, match ? match.index : text.length);
		// The final line of a prefix may be cut mid-word; a whole one always has
		// its break, so only an unterminated last line is skipped.
		if (match === null && options.truncated && line !== '') {
			break;
		}
		if (PROCEDURE_HEADER_RE.test(line)) {
			return { complete: true };
		}
		const comment = commentTextOf(line);
		const tag = comment === undefined ? null : FOLDER_TAG_RE.exec(comment);
		if (comment !== undefined && tag) {
			const folder = normalizeFolderPath(folderArgument(comment.slice(tag.index + tag[0].length)));
			return folder ? { folder, complete: true } : { complete: true };
		}
		if (match === null) {
			break;
		}
		start = LINE_BREAK_RE.lastIndex;
	}
	return { complete: false };
}

/**
 * The path as the tree stores it: segments trimmed, empty segments dropped,
 * joined by dots. Nothing left means the module names no folder.
 */
export function normalizeFolderPath(raw: string): string {
	return raw
		.split('.')
		.map((segment) => segment.trim())
		.filter((segment) => segment.length > 0)
		.join('.');
}

/**
 * The comment part of a line, or undefined when the line carries none. The
 * scan steps over string literals so the apostrophe in `Const S = "it's"`
 * does not open a comment.
 */
function commentTextOf(line: string): string | undefined {
	let inString = false;
	// `Rem` is a comment only as a whole word starting the statement, so the
	// test is worth making once - not at every "r" in a long declaration.
	let atStatementStart = true;
	for (let i = 0; i < line.length; i += 1) {
		const ch = line[i];
		if (ch === '"') {
			inString = !inString;
			atStatementStart = false;
			continue;
		}
		if (inString) {
			continue;
		}
		if (ch === "'") {
			return line.slice(i + 1);
		}
		if (atStatementStart) {
			if (ch === ' ' || ch === '\t') {
				continue;
			}
			if ((ch === 'R' || ch === 'r') && REM_COMMENT_RE.test(line.slice(i))) {
				return line.slice(i + 3);
			}
			atStatementStart = false;
		}
	}
	return undefined;
}

/**
 * The argument after the tag, in any of the written forms. Reading forward
 * (rather than trimming brackets off both ends) keeps prose after the
 * annotation out of the folder name.
 */
function folderArgument(rest: string): string {
	let i = 0;
	const skipSpace = (): void => {
		while (i < rest.length && (rest[i] === ' ' || rest[i] === '\t')) { i += 1; }
	};
	skipSpace();
	if (rest[i] === '(') {
		i += 1;
		skipSpace();
	}
	if (rest[i] === '"') {
		i += 1;
		const close = rest.indexOf('"', i);
		return close === -1 ? rest.slice(i) : rest.slice(i, close);
	}
	const close = rest.indexOf(')', i);
	return close === -1 ? rest.slice(i) : rest.slice(i, close);
}
