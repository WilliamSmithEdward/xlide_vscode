// The spellings the VBE gives conditional-compilation names, for the
// formatter's identifier casing.
//
// The VBE keeps one spelling per name for the whole project, and a name a
// `#Const` line or the project's Conditional Compilation Arguments declare is
// in it like any other: `#If debugmode Then` prints as the `#Const DebugMode`
// that declared it, in that module and every other one, and so does a code
// line that names it. Where one name is declared more than once, the last
// declaration wins, in the order the VBE loads them: the project's arguments,
// then each module in project order, top to bottom. A `#Const` inside a branch
// that is not compiled counts too. With no declaration, VBA's own compiler
// constants print as VBA spells them: `#If vba7 And Not mac Then` becomes
// `#If VBA7 And Not Mac Then`. Measured in Excel, Word, PowerPoint and Access
// 16.0 by opening files with the modules written in.

import { tokenize } from '../lexer/tokenize';
import { tokenWord } from '../lexer/tokenHelpers';
import {
	parseProjectConditionalConstants,
	VBA_COMPILER_CONSTANT_NAMES,
} from '../conditional/conditionalCompilation';

const COMPILER_CONSTANTS = new Map(VBA_COMPILER_CONSTANT_NAMES.map((name) => [name.toLowerCase(), name]));

let declarationScans = 0;

/** How many module sources {@link conditionalConstantDeclarations} has tokenized. */
export function conditionalDeclarationScanCountForTests(): number {
	return declarationScans;
}

/**
 * The names a module's `#Const` lines spell, in source order, a bracketed
 * name without its brackets (`#Const [Flag]` spells `Flag`, which is how the
 * VBE shows it). The VBE takes the spelling from a `#Const` that opens its
 * line, with a continuation after the `#` too, and also from one with `Rem` or
 * a `:` after its value, although such a line defines nothing (#1325, #1330);
 * not from one after a statement and a `:` (`x = 1: #Const Flag = 1`).
 */
export function conditionalConstantDeclarations(source: string): string[] {
	// Most modules have no #Const; those are not tokenized at all. `_` lets a
	// continuation sit between the `#` and `Const`.
	if (!/#[\s_]*const\b/i.test(source)) {
		return [];
	}
	declarationScans++;
	// Not the shared token cache: the caller keeps the answer per source, and
	// a whole project's modules would evict what the editor is using.
	const tokens = tokenize(source);
	const names: string[] = [];
	for (let i = 0; i + 2 < tokens.length; i++) {
		if (tokens[i].kind !== 'directive' || tokenWord(tokens[i + 1]) !== 'const'
			|| (i > 0 && tokens[i - 1].kind !== 'newline')) {
			continue;
		}
		const name = tokens[i + 2];
		if (name.kind === 'identifier') {
			names.push(name.rawText);
		} else if (name.kind === 'bracketedIdentifier' && name.rawText.length > 2 && name.rawText.endsWith(']')) {
			names.push(name.rawText.slice(1, -1));
		}
	}
	return names;
}

/**
 * The spelling of a conditional-compilation name, or undefined when neither
 * the project's arguments, the modules' `#Const` declarations (each module's
 * {@link conditionalConstantDeclarations}, in project order) nor VBA's own
 * compiler constants name it.
 */
export function conditionalNameCase(
	moduleDeclarations: Iterable<readonly string[]>,
	projectConstantsRaw?: string,
): (name: string) => string | undefined {
	const declared = new Map<string, string>();
	for (const name of Object.keys(parseProjectConditionalConstants(projectConstantsRaw))) {
		declared.set(name.toLowerCase(), name);
	}
	for (const names of moduleDeclarations) {
		for (const name of names) {
			declared.set(name.toLowerCase(), name);
		}
	}
	return (name) => {
		const lower = name.toLowerCase();
		return declared.get(lower) ?? COMPILER_CONSTANTS.get(lower);
	};
}
