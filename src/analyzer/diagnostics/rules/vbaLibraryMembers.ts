// Rule: names after `VBA.` that the VBA library does not have (issue #369),
// and after `Excel.` in an Excel project (issue #305).
// Measured in Excel 16.0, compiled with Debug > Compile:
//
//  - `VBA.Nosuch`, `VBA.Strings.Nosuch`, `VBA.VbMsgBoxResult.vbNosuch` and
//    `VBA.Global.Left$(...)` are "Method or data member not found": the
//    library, its modules, its enums and its hidden Global class are closed
//    (see runtime/vbaLibraryNames.ts, generated from VBE7.DLL).
//  - `VBA.Err.LastDllError = 5` is "Can't assign to read-only property", as
//    the unqualified `Err.LastDllError = 5` is. Err's other members are not
//    judged: `VBA.Err.Nosuch` compiles.
//
// A project that declares a name VBA of its own is not judged.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode } from '../../parser/nodes';
import { EXCEL_LIBRARY_NAMES } from '../../host/excelLibraryNames';
import { VBA_ERR_READ_ONLY, VBA_LIBRARY_CONTAINERS, VBA_LIBRARY_NAMES } from '../../runtime/vbaLibraryNames';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import {
	activeModuleMembers,
	firstExecutableTokenIndex,
	forEachStatement,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

export function checkVbaLibraryMembers(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	hostName?: string,
): void {
	const shadows = (lower: string): boolean => (symbols.root.children ?? []).some((symbol) => symbol.name.toLowerCase() === lower || (symbol.children ?? []).some((child) => child.name.toLowerCase() === lower))
		|| (projectVisibleSymbols ?? []).some((symbol) => symbol.name.toLowerCase() === lower || symbol.moduleName.toLowerCase() === lower);
	if (hostName === undefined || hostName === 'Excel') {
		if (!shadows('excel')) {
			checkExcelLibraryNames(source, mod, activity, push);
		}
	}
	if (shadows('vba')) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const at = (k: number) => ({ start: span.start + toks[k].start, end: span.start + toks[k].end });
				for (let i = 0; i + 2 < toks.length; i++) {
					if (tokenText(toks[i]) !== 'vba' || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.') {
						continue;
					}
					const first = nameAt(toks, i + 2);
					if (!first) {
						continue;
					}
					if (!VBA_LIBRARY_NAMES.has(first.lower)) {
						push('memberNotFound', missing(first, 'the VBA library', VBA_LIBRARY_NAMES), at(i + 2));
						continue;
					}
					if (toks[first.next]?.rawText !== '.') {
						continue;
					}
					const second = nameAt(toks, first.next + 1);
					if (!second) {
						continue;
					}
					const container = VBA_LIBRARY_CONTAINERS.get(first.lower);
					if (container && !container.has(second.lower)) {
						push('memberNotFound', missing(second, `VBA.${first.shown}`, container), at(first.next + 1));
						continue;
					}
					// `VBA.Err.LastDllError = 5`, the assignment's target.
					const head = firstExecutableTokenIndex(toks);
					const startsTarget = i === head || (i === head + 1 && ['let', 'set'].includes(tokenText(toks[head])));
					if (first.lower === 'err' && startsTarget && VBA_ERR_READ_ONLY.has(second.lower) && toks[second.next]?.rawText === '=') {
						push('readonlyMemberAssignment', `Cannot assign to read-only property 'VBA.Err.${second.shown}'. This is a VBE compile error: Can't assign to read-only property.`, at(first.next + 1));
					}
				}
			}
		}, activity);
	}
}

/**
 * `Excel.Nope`, `Excel.Version`: a name after `Excel.` that is no type,
 * enum constant or Global member of the library (issue #305, measured in
 * Excel 16.0). A name after As or New is a type, which the VBE refuses
 * otherwise, and a member of what `Excel.X` gives is not judged here.
 */
function checkExcelLibraryNames(source: string, mod: ModuleNode, activity: ConditionalActivityTracker | undefined, push: PushFn): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				for (let i = 0; i + 2 < toks.length; i++) {
					if (tokenText(toks[i]) !== 'excel' || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.'
						|| ['as', 'new'].includes(tokenText(toks[i - 1]) ?? '')) {
						continue;
					}
					const name = tokenName(toks[i + 2]);
					if (name && !EXCEL_LIBRARY_NAMES.has(name.toLowerCase())) {
						push('memberNotFound', `'${name}' is not a member of the Excel library. This is a VBE compile error: Method or data member not found.`, { start: span.start + toks[i + 2].start, end: span.start + toks[i + 2].end });
					}
				}
			}
		}, activity);
	}
}

/**
 * Why a name is refused: none of that name, or a `$` form of a name that has
 * none, `VBA.Asc$` (measured in Excel 16.0).
 */
function missing(name: { lower: string; shown: string }, where: string, names: ReadonlySet<string>): string {
	if (name.lower.endsWith('$') && names.has(name.lower.slice(0, -1))) {
		return `'${name.shown}' has no String form with $ in ${where}. This is a VBE compile error: Type-declaration character does not match declared data type.`;
	}
	return `'${name.shown}' is not a member of ${where}. This is a VBE compile error: Method or data member not found.`;
}

/** The name at `k`, lowercased, with a `$` glued to it, and the index after it. */
function nameAt(toks: readonly VbaToken[], k: number): { lower: string; shown: string; next: number } | undefined {
	const name = tokenName(toks[k]);
	if (!name) {
		return undefined;
	}
	const dollar = toks[k + 1]?.rawText === '$' && toks[k + 1].start === toks[k].end;
	return { lower: name.toLowerCase() + (dollar ? '$' : ''), shown: toks[k].rawText + (dollar ? '$' : ''), next: k + (dollar ? 2 : 1) };
}
