// The file numbers a module's Open statements name (issue #419). A file
// statement on a number no Open in the project names raises 52, "Bad file
// name or number", wherever it runs; one Open whose number is a variable or
// FreeFile may open any number, and then none is judged.

import { tokenWord } from '../lexer/tokenHelpers';
import { tokenizeCached } from '../lexer/tokenize';

/** What a module's, or a project's, Open statements name. */
export interface OpenedFileNumbers {
	/** An Open whose number is not a literal: it may open any. */
	any: boolean;
	/** The literal numbers Opens name. */
	numbers: ReadonlySet<number>;
}

/**
 * The numbers `Open ... As #n` names in a module's source: a statement of
 * its own, after a colon or a line number, or a one-line If's arm.
 */
export function openedFileNumbersIn(source: string): OpenedFileNumbers {
	const toks = tokenizeCached(source).filter((tok) => tok.kind !== 'comment');
	const numbers = new Set<number>();
	let any = false;
	for (let i = 0; i < toks.length; i++) {
		if (tokenWord(toks[i]) !== 'open') {
			continue;
		}
		const before = toks[i - 1];
		const lineNumber = before?.kind === 'integerLiteral' && (i < 2 || toks[i - 2].kind === 'newline');
		const starts = before === undefined || before.kind === 'newline' || before.kind === 'colon' || lineNumber
			|| tokenWord(before) === 'then' || tokenWord(before) === 'else';
		if (!starts) {
			continue;
		}
		for (let j = i + 1; j < toks.length && toks[j].kind !== 'newline' && toks[j].kind !== 'colon'; j++) {
			if (tokenWord(toks[j]) !== 'as') {
				continue;
			}
			const number = toks[j + 1]?.rawText === '#' ? toks[j + 2] : toks[j + 1];
			if (number?.kind === 'integerLiteral' && /^\d+$/.test(number.rawText)) {
				numbers.add(Number(number.rawText));
			} else {
				any = true;
			}
			break;
		}
	}
	return { any, numbers };
}

/** Both modules' Opens, or a project's and one module's current text. */
export function mergeOpenedFileNumbers(parts: readonly OpenedFileNumbers[]): OpenedFileNumbers {
	return { any: parts.some((part) => part.any), numbers: new Set(parts.flatMap((part) => [...part.numbers])) };
}
