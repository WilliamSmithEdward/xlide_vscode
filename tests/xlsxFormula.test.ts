import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import {
	FormulaError,
	formulaForDisplay,
	formulaForFile,
	shiftFormula,
	type FormulaContext,
} from '../src/vba/xlsxFormula';

// The two recorded fixtures hold what Excel 16 itself did: which formulas its
// parser accepted, and the text it stored for formulas typed into cells.
const FIXTURES = path.join(__dirname, 'fixtures');

interface RecordedContext {
	sheets: string[];
	names: string[];
	tables: Record<string, string[]>;
}

function contextOf(recorded: RecordedContext): FormulaContext {
	return {
		sheets: new Set(recorded.sheets),
		names: new Set(recorded.names),
		tables: new Map(Object.entries(recorded.tables).map(([name, columns]) => [name, new Set(columns)])),
	};
}

const context = contextOf({ sheets: ['SHEET1', 'DATA'], names: ['TBL', 'RATE'], tables: { TBL: ['Q1', 'VAL'] } });
const stored = (formula: string): string => formulaForFile(formula, context);
const refused = (formula: string): string => {
	try {
		stored(formula);
	} catch (e) {
		expect(e).toBeInstanceOf(FormulaError);
		return (e as Error).message;
	}
	throw new Error(`${formula} was accepted`);
};

describe('shiftFormula', () => {
	it('moves relative references and leaves anchored parts', () => {
		expect(shiftFormula('A2*$B$2+B$2+$A3', 2, 1)).toBe('B4*$B$2+C$2+$A5');
	});

	it('does not touch function names that look like cells', () => {
		expect(shiftFormula('LOG10(A2)+DEC2BIN(A2)&A2', 1, 0)).toBe('LOG10(A3)+DEC2BIN(A3)&A3');
	});

	it('moves whole-row and whole-column references', () => {
		expect(shiftFormula('SUM(2:2)+A2', 1, 0)).toBe('SUM(3:3)+A3');
		expect(shiftFormula('SUM(A:A)+A$1', 0, 2)).toBe('SUM(C:C)+C$1');
	});

	it('leaves strings, quoted sheet names and table columns alone', () => {
		expect(shiftFormula('"A1"&\'Q1\'!A2&SUM(Tbl[Q1])', 1, 0)).toBe('"A1"&\'Q1\'!A3&SUM(Tbl[Q1])');
	});

	it('gives #REF! for a reference moved off the sheet', () => {
		expect(shiftFormula('A1+B2', -1, 0)).toBe('#REF!+B1');
	});
});

describe('formulaForDisplay', () => {
	it('drops the prefixes a file stores', () => {
		expect(formulaForDisplay('_xlfn.XLOOKUP(2,A2:A4,A2:A4)+_xlfn._xlws.SORT(A2:A4)')).toBe('XLOOKUP(2,A2:A4,A2:A4)+SORT(A2:A4)');
		expect(formulaForDisplay('_xlfn.LET(_xlpm.x,2,_xlpm.x*3)')).toBe('LET(x,2,x*3)');
		expect(formulaForDisplay('_xlfn.LAMBDA(_xlpm.x,_xlop.y,_xlpm.x)(1)')).toBe('LAMBDA(x,[y],x)(1)');
		expect(formulaForDisplay('_xlfn.GROUPBY(A2:A4,B2:B4,_xleta.SUM)')).toBe('GROUPBY(A2:A4,B2:B4,SUM)');
	});

	it('shows the spill, intersection and trim operators', () => {
		expect(formulaForDisplay('SUM(_xlfn.ANCHORARRAY(F2))')).toBe('SUM(F2#)');
		expect(formulaForDisplay('_xlfn.SINGLE(A2:A4)')).toBe('@A2:A4');
		expect(formulaForDisplay('_xlfn.SINGLE((A2,A4))')).toBe('@(A2,A4)');
		expect(formulaForDisplay('SUM(_xlfn._TRO_TRAILING(A2:A4))')).toBe('SUM(A2:.A4)');
	});

	it('shows a this-row table reference with @', () => {
		expect(formulaForDisplay('Tbl[[#This Row],[Val]]*2')).toBe('Tbl[@Val]*2');
	});
});

describe('formulaForFile', () => {
	it('prefixes newer functions', () => {
		expect(stored('XLOOKUP(2,A2:A4,A2:A4)')).toBe('_xlfn.XLOOKUP(2,A2:A4,A2:A4)');
		expect(stored('sort(A2:A4)')).toBe('_xlfn._xlws.SORT(A2:A4)');
		expect(stored('SUM(A2:A4)')).toBe('SUM(A2:A4)');
		expect(stored('MyUdf(A2)')).toBe('MyUdf(A2)');
	});

	it('names LET and LAMBDA parameters as Excel stores them', () => {
		expect(stored('LET(x,2,y,x+1,x*y)')).toBe('_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)');
		expect(stored('LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1)'))
			.toBe('_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))(1)');
	});

	it('passes a built-in by name as an eta lambda, unless a workbook name shadows it', () => {
		expect(stored('GROUPBY(A2:A4,B2:B4,SUM)')).toBe('_xlfn.GROUPBY(A2:A4,B2:B4,_xleta.SUM)');
		expect(stored('Rate*2')).toBe('Rate*2');
	});

	it('spells out the spill, intersection and trim operators', () => {
		expect(stored('SUM(F2#)')).toBe('SUM(_xlfn.ANCHORARRAY(F2))');
		expect(stored('@A2,A4')).toBe('_xlfn.SINGLE(A2),A4');
		expect(stored('-@A2')).toBe('-_xlfn.SINGLE(A2)');
		expect(stored('SUM(A2.:.A4)')).toBe('SUM(_xlfn._TRO_ALL(A2:A4))');
		expect(stored('SUM(A2 : A4)')).toBe('SUM(A2:A4)');
	});

	it('checks and spells out table references', () => {
		expect(stored('Tbl[@Val]*2')).toBe('Tbl[[#This Row],[Val]]*2');
		expect(stored('SUM(Tbl[[Q1]:[Val]])')).toBe('SUM(Tbl[[Q1]:[Val]])');
		expect(refused('SUM(Tbl[Missing])')).toMatch(/no column named Missing/);
		expect(refused('SUM(Other[Val])')).toMatch(/not a table/);
		expect(refused('[@Val]')).toMatch(/needs its table name/);
	});

	it('closes a lone call left open, as Excel does, and nothing else', () => {
		expect(stored('SUM(A2,A3')).toBe('SUM(A2,A3)');
		expect(refused('SUM(ABS(A2')).toMatch(/never closed/);
		expect(refused('A2+SUM(A2')).toMatch(/never closed/);
	});

	it('refuses what Excel refuses', () => {
		expect(refused('SUM()')).toMatch(/SUM takes 1 to 255 arguments, not 0/);
		expect(refused('IF(1)')).toMatch(/IF takes 2 to 3 arguments/);
		expect(refused('SUMIFS(A2:A4,A2:A4)')).toMatch(/in steps of 2/);
		expect(refused('SUMIF(1,">0")')).toMatch(/argument 1 of SUMIF must be a reference/);
		expect(refused('1*/2')).toMatch(/is not a value/);
		expect(refused('SUM(A2;A3)')).toMatch(/separated by commas/);
		expect(refused('A2+')).toMatch(/ends where a value is expected/);
		expect(refused('SUM(2#)')).toMatch(/spill operator/);
		expect(refused('2 3')).toMatch(/intersection/);
		expect(refused('(1,2)')).toMatch(/union/);
		expect(refused('LAMBDA()')).toMatch(/LAMBDA needs a calculation/);
		expect(refused('LET(x,1,x,2,x)')).toMatch(/value twice/);
		expect(refused('LET(a.b,1,a.b)')).toMatch(/letters, digits and underscores/);
		expect(refused('{1,2;3}')).toMatch(/same length/);
		expect(refused('"abc')).toMatch(/string is not closed/);
		expect(refused('SPLIT("a,b",",")')).toMatch(/SPLIT is not a function Excel accepts/);
		expect(refused('C1A(1)')).toMatch(/C1A is not a function/);
		expect(refused('Sheet9!A1')).toMatch(/'Sheet9' is not a sheet/);
		expect(refused('[Book2.xlsx]Sheet1!A1')).toMatch(/another workbook/);
		expect(refused('   ')).toMatch(/empty/);
		expect(refused(`${'1+'.repeat(4100)}1`)).toMatch(/limit of 8192/);
	});

	it('never accepts a formula Excel rejected', () => {
		const recorded = JSON.parse(fs.readFileSync(path.join(FIXTURES, 'excelFormulaVerdicts.json'), 'utf8')) as {
			context: RecordedContext;
			verdicts: Array<[string, boolean]>;
		};
		const recordedContext = contextOf(recorded.context);
		const dangerous: string[] = [];
		let agreed = 0;
		for (const [formula, excel] of recorded.verdicts) {
			let ours = true;
			try { formulaForFile(formula, recordedContext); } catch { ours = false; }
			if (ours && !excel) { dangerous.push(formula); }
			if (ours === excel) { agreed++; }
		}
		expect(dangerous).toEqual([]);
		// The rest of the disagreements are refusals on the safe side, such as a
		// sheet that does not exist, which Excel takes as a link to a file.
		expect(agreed).toBeGreaterThan(recorded.verdicts.length * 0.97);
	});

	it('stores and shows formulas as Excel does', () => {
		const groups = JSON.parse(fs.readFileSync(path.join(FIXTURES, 'excelStoredFormulas.json'), 'utf8')) as Array<{
			context: RecordedContext;
			entries: Array<{ typed: string; stored: string; shown: string }>;
		}>;
		const differences: string[] = [];
		for (const group of groups) {
			const groupContext = contextOf(group.context);
			for (const entry of group.entries) {
				const file = formulaForFile(entry.typed, groupContext);
				if (file !== entry.stored) { differences.push(`${entry.typed}: stored ${file}, Excel ${entry.stored}`); }
				const shown = formulaForDisplay(entry.stored);
				if (shown !== entry.shown) { differences.push(`${entry.stored}: shown ${shown}, Excel ${entry.shown}`); }
			}
		}
		expect(differences).toEqual([]);
	});
});
