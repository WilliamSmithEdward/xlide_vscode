// The spelling the VBE gives conditional-compilation names (issue #1497).
//
// The VBE keeps one spelling per name for the whole project. For the names
// conditional compilation declares, measured in Excel, Word, PowerPoint and
// Access 16.0 (build 20527) by opening files XLIDE wrote
// (tests/office/directiveCasing.office.ts):
//
// - VBA's compiler constants print as VBA spells them, on directive and code
//   lines alike: `#If vba7` shows `#If VBA7`, `x = win64` shows `x = Win64`.
//   `TWINBASIC` is not one of them, and stays as written.
// - A name the project's Conditional Compilation Arguments or a `#Const` line
//   declare prints as its last declaration in load order: the arguments,
//   then each module in project order, top to bottom. Every `#Const` counts,
//   in an uncompiled branch or inside a procedure too, and it overrides a
//   compiler constant's spelling: `#Const WIN64 = 0` makes every `win64`
//   print `WIN64`.
//
// An ordinary declaration of the same name takes over as well (`Public mac
// As Long` keeps `mac`), which the caller gives precedence to.

import { tokenize } from '../lexer/tokenize';
import { tokenWord } from '../lexer/tokenHelpers';

/** VBA's own compiler constants, spelled as the VBE prints them. */
export const VBA_COMPILER_CONSTANTS: readonly string[] = ['VBA6', 'VBA7', 'Win16', 'Win32', 'Win64', 'Mac'];

/** Cheap test for a `#Const` anywhere, before the module is lexed. */
const MAY_DECLARE = /#[ \t]*(?:_[ \t]*\r?\n[ \t]*)*Const\b/i;

/**
 * The names a module's `#Const` lines declare, top to bottom. A `#Const`
 * counts only where a directive can start: at the start of a logical line,
 * not after a statement and a `:`.
 */
export function conditionalConstNames(source: string): string[] {
	if (!MAY_DECLARE.test(source)) {
		return [];
	}
	const names: string[] = [];
	const tokens = tokenize(source);
	for (let i = 0; i + 2 < tokens.length; i++) {
		if (tokens[i].kind !== 'directive' || tokenWord(tokens[i + 1]) !== 'const') {
			continue;
		}
		const previous = tokens[i - 1];
		if (previous && previous.kind !== 'newline') {
			continue;
		}
		const name = tokens[i + 2];
		if (name.kind === 'identifier') {
			names.push(name.rawText);
		} else if (name.kind === 'bracketedIdentifier') {
			names.push(name.rawText.slice(1, -1));
		}
	}
	return names;
}

/** The names `Name = Value : Name = Value` declares, in order. */
export function projectArgumentNames(raw: string | undefined): string[] {
	const names: string[] = [];
	for (const entry of (raw ?? '').split(':')) {
		const eq = entry.indexOf('=');
		const name = eq < 0 ? '' : entry.slice(0, eq).trim();
		if (/^\p{L}[\p{L}\p{N}_]*$/u.test(name)) {
			names.push(name);
		}
	}
	return names;
}

/**
 * Lowercased name -> its VBE spelling, from the compiler constants, then the
 * project's arguments, then each module's `#Const` names in load order. A
 * later declaration replaces an earlier one's spelling.
 */
export function conditionalNameSpellings(
	projectArguments: string | undefined,
	moduleConstNames: Iterable<readonly string[]>,
): Map<string, string> {
	const spellings = new Map<string, string>();
	const declare = (name: string): void => {
		spellings.set(name.toLowerCase(), name);
	};
	VBA_COMPILER_CONSTANTS.forEach(declare);
	projectArgumentNames(projectArguments).forEach(declare);
	for (const names of moduleConstNames) {
		names.forEach(declare);
	}
	return spellings;
}
