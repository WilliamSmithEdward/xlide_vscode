// Rule: host properties set to a literal the host refuses (issue #204).
//
// The type libraries give these properties a type, not a range, so the
// ranges are a table. Every range below is one Excel, Word or PowerPoint 16.0
// was seen to refuse, through pyVBAharness on 2026-09-29, beside values it
// took: only values inside a refused range are reported, never a value the
// table merely does not know. `Font.Size = 409.5` runs in Excel and 409.6
// raises; nothing between is claimed.
//
// The enum-typed ones (issue #244) were swept from -100 to 999 in Excel
// 16.0: the refused ranges are the runs of that sweep that raised, so a
// constant such as xlCenter (-4108) or xlPatternLinearGradient (4000),
// outside every run, is never claimed. Calculation, CutCopyMode and
// ReferenceStyle took every value of the sweep, and are not here.
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
		['underline', { runs: '1 to 5, or an xlUnderlineStyle constant', refused: [{ from: -100, to: 0 }, { from: 6, to: 999 }], error: EXCEL_1004('Underline', 'Font') }],
	])],
	['Excel.Interior', new Map([
		['colorindex', { runs: '1 to 56, or an xlColorIndex constant', refused: [{ from: 57 }], error: SUBSCRIPT }],
		['pattern', { runs: '-1 to 18, or an xlPattern constant', refused: [{ from: -100, to: -2 }, { from: 19, to: 999 }], error: SUBSCRIPT }],
	])],
	['Excel.Border', new Map([
		['linestyle', { runs: '0 to 13, or an xlLineStyle constant', refused: [{ from: -100, to: -1 }, { from: 14, to: 999 }], error: EXCEL_1004('LineStyle', 'Border') }],
		['weight', { runs: '1 to 4, or an xlBorderWeight constant', refused: [{ from: -100, to: 0 }, { from: 5, to: 999 }], error: EXCEL_1004('Weight', 'Border') }],
	])],
	['Excel.PageSetup', new Map([
		['orientation', { runs: 'xlPortrait (1) or xlLandscape (2)', refused: [{ from: -100, to: 0 }, { from: 3, to: 999 }], error: EXCEL_1004('Orientation', 'PageSetup') }],
	])],
	['Excel.Worksheet', new Map([
		['visible', { runs: 'an xlSheetVisibility constant: -1, 0 or 2', refused: [{ from: -100, to: -2 }, { from: 3, to: 999 }], error: EXCEL_1004('Visible', 'Worksheet') }],
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
		['horizontalalignment', { runs: '1 to 8, or an xlHAlign constant', refused: [{ from: -100, to: 0 }, { from: 9, to: 999 }], error: EXCEL_1004('HorizontalAlignment', 'Range') }],
		['verticalalignment', { runs: '1 to 5, or an xlVAlign constant', refused: [{ from: -100, to: 0 }, { from: 6, to: 999 }], error: EXCEL_1004('VerticalAlignment', 'Range') }],
	])],
	['Excel.Window', new Map([
		// -1 is True, which fits the selection.
		['zoom', { runs: '10 to 400, or True', refused: [{ to: 9 }, { from: 401 }], allowed: [-1], error: EXCEL_1004('Zoom', 'Window') }],
		['windowstate', { runs: '1 to 3, or an xlWindowState constant', refused: [{ from: -100, to: 0 }, { from: 4, to: 999 }], error: EXCEL_1004('WindowState', 'Window') }],
	])],
	['Word.Font', new Map([
		['size', { runs: '1 to 1638', refused: [{ to: 0.5 }, { from: 1638.5 }], error: WORD_RANGE }],
		// Swept from -100 to 999 in Word 16.0 (issue #245): the wdUnderline
		// values are scattered, and each gap between them is refused.
		['underline', {
			runs: 'a wdUnderline constant: -1 to 4, 6, 7, 9 to 11, 20, 23, 25 to 27, 39, 43 or 55',
			refused: [{ from: -100, to: -2 }, { from: 5, to: 5 }, { from: 8, to: 8 }, { from: 12, to: 19 }, { from: 21, to: 22 }, { from: 24, to: 24 },
				{ from: 28, to: 38 }, { from: 40, to: 42 }, { from: 44, to: 54 }, { from: 56, to: 999 }],
			error: WORD_RANGE,
		}],
	])],
	['Word.Zoom', new Map([
		['percentage', { runs: '10 to 500', refused: [{ to: 9 }, { from: 501 }], error: WORD_RANGE }],
	])],
	['Word.Paragraph', new Map([
		['alignment', { runs: '0 to 9', refused: [{ from: -100, to: -1 }, { from: 10, to: 999 }], error: { number: '5148', text: 'The number must be between 0 and 9' } }],
		['linespacingrule', { runs: '0 to 5', refused: [{ from: -100, to: -1 }, { from: 6, to: 999 }], error: { number: '5148', text: 'The number must be between 0 and 5' } }],
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
