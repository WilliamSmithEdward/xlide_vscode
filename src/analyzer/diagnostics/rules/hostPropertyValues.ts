// Rule: host properties set to a literal the host refuses (issue #204).
//
// The type libraries give these properties a type, not a range, so the
// ranges are a table. Every range below is one Excel, Word or PowerPoint 16.0
// was seen to refuse, through pyVBAharness on 2026-09-29, beside values it
// took: only values inside a refused range are reported, never a value the
// table merely does not know. `Font.Size = 409.5` runs in Excel and 409.6
// raises; nothing between is claimed.
//
// Only a numeric literal is read, optionally signed. `ActiveWindow.Zoom =
// False` runs where `Zoom = 0` raises, and a named constant such as
// xlVertical is its own value, not a number the table can place.

import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { MemberCompletion } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';

interface HostValueLimit {
	/** The range that runs, as the message states it. */
	runs: string;
	/** Inclusive ranges the host refuses; an open end is unbounded. */
	refused: ReadonlyArray<{ from?: number; to?: number }>;
	/** Values inside a refused range that run: the host's named constants. */
	allowed?: readonly number[];
	error: { number: string; text: string };
}

const EXCEL_1004 = (property: string, owner: string) => ({ number: '1004', text: `Unable to set the ${property} property of the ${owner} class` });
const SUBSCRIPT = { number: '9', text: 'Subscript out of range' };
const WORD_RANGE = { number: '5843', text: 'One of the values passed to this method or property is out of range' };

/** By qualified owner type, then lower-cased property name. */
const LIMITS: ReadonlyMap<string, ReadonlyMap<string, HostValueLimit>> = new Map([
	['Excel.Font', new Map([
		['size', { runs: '1 to 409.5', refused: [{ to: 0.5 }, { from: 409.6 }], error: EXCEL_1004('Size', 'Font') }],
	])],
	['Excel.Interior', new Map([
		['colorindex', { runs: '1 to 56, or an xlColorIndex constant', refused: [{ from: 57 }], error: SUBSCRIPT }],
	])],
	['Excel.Tab', new Map([
		['colorindex', { runs: '1 to 56, or xlColorIndexNone', refused: [{ from: -1, to: 0 }, { from: 57 }], error: SUBSCRIPT }],
	])],
	['Excel.Range', new Map([
		['rowheight', { runs: '0 to 409.5', refused: [{ to: -0.5 }, { from: 409.75 }], error: EXCEL_1004('RowHeight', 'Range') }],
		['columnwidth', { runs: '0 to 255', refused: [{ to: -0.5 }, { from: 255.5 }], error: EXCEL_1004('ColumnWidth', 'Range') }],
		// xlHorizontal, xlVertical, xlUpward and xlDownward are the constants.
		['orientation', { runs: '-90 to 90, or an xlOrientation constant', refused: [{ to: -91 }, { from: 91 }], allowed: [-4128, -4166, -4171, -4170], error: EXCEL_1004('Orientation', 'Range') }],
		['indentlevel', { runs: 'up to 250', refused: [{ to: -16 }, { from: 251 }], error: EXCEL_1004('IndentLevel', 'Range') }],
	])],
	['Excel.Window', new Map([
		// -1 is True, which fits the selection.
		['zoom', { runs: '10 to 400, or True', refused: [{ to: 9 }, { from: 401 }], allowed: [-1], error: EXCEL_1004('Zoom', 'Window') }],
	])],
	['Word.Font', new Map([
		['size', { runs: '1 to 1638', refused: [{ to: 0.5 }, { from: 1638.5 }], error: WORD_RANGE }],
	])],
	['Word.Zoom', new Map([
		['percentage', { runs: '10 to 500', refused: [{ to: 9 }, { from: 501 }], error: WORD_RANGE }],
	])],
	['Word.Paragraph', new Map([
		['leftindent', { runs: '-1584 to 1584 points', refused: [{ to: -1585 }, { from: 1585 }], error: { number: '5149', text: 'The measurement must be between -1584 pt and 1584 pt' } }],
	])],
	['PowerPoint.Font', new Map([
		['size', { runs: '1 to 4000', refused: [{ to: 0 }, { from: 4000.25 }], error: { number: '-2147024809', text: 'The specified value is out of range' } }],
	])],
]);

/** The number a value's tokens spell: a numeric literal, optionally signed. */
export function signedNumericLiteral(tokens: readonly VbaToken[]): number | undefined {
	const toks = tokens.filter((tok) => tok.kind !== 'comment');
	const signed = toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+');
	const literal = signed ? toks[1] : toks.length === 1 ? toks[0] : undefined;
	let value: number | undefined;
	if (literal?.kind === 'integerLiteral') {
		value = parseVbaIntegerLiteral(literal.rawText);
	} else if (literal?.kind === 'floatLiteral') {
		value = Number(literal.rawText.replace(/[!#@]$/, ''));
	}
	if (value === undefined || !Number.isFinite(value)) {
		return undefined;
	}
	return signed && toks[0].rawText === '-' ? -value : value;
}

/** The message for a host property Let the host refuses, or undefined. */
export function hostPropertyValueProblem(target: MemberCompletion, valueTokens: readonly VbaToken[]): string | undefined {
	const limit = LIMITS.get(target.owner)?.get(target.name.toLowerCase());
	const value = limit ? signedNumericLiteral(valueTokens) : undefined;
	if (!limit || value === undefined || limit.allowed?.includes(value)) {
		return undefined;
	}
	const refused = limit.refused.some((range) => (range.from === undefined || value >= range.from) && (range.to === undefined || value <= range.to));
	if (!refused) {
		return undefined;
	}
	const bare = target.owner.slice(target.owner.indexOf('.') + 1);
	return `${bare}.${target.name} takes ${limit.runs}; ${value} is outside that. This will raise Run-time error '${limit.error.number}': ${limit.error.text}.`;
}
