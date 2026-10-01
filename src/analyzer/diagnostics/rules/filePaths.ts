// Rule: a file statement or function given an empty path (issue #262).
//
// Each was measured in Excel 16.0 (build 20326, 2026-10-01): `Open ""` raises
// 75, `FileLen("")` 53 and `MkDir ""` 76, and some take a path of spaces the
// same way while Dir("") runs and Dir(" ") raises 53. Only the pairs measured
// are reported. The path is a string literal, or a String local or Const the
// statement is known to see holding one: a String never assigned holds "".

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import {
	knownLocalLiteralValuesAt,
	runtimeCallableSourceShadowed,
	sourceNameScopeFor,
	stringConstantsInScope,
	stringLiteralValue,
} from '../typeInference';
import {
	matchParenFrom,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';
import { isBareOrVbaQualifiedIntrinsicCall } from './shared';

/** The error each statement or function raises for "" and for a path of spaces, where measured. */
const ERRORS: Readonly<Record<string, { empty?: number; blank?: number; display: string }>> = {
	open: { empty: 75, blank: 53, display: 'Open' },
	mkdir: { empty: 76, blank: 76, display: 'MkDir' },
	chdir: { empty: 76, blank: 76, display: 'ChDir' },
	rmdir: { empty: 76, display: 'RmDir' },
	kill: { empty: 53, display: 'Kill' },
	setattr: { empty: 53, display: 'SetAttr' },
	filecopy: { empty: 75, display: 'FileCopy' },
	name: { empty: 75, display: 'Name' },
	filelen: { empty: 53, blank: 53, display: 'FileLen' },
	filedatetime: { empty: 53, display: 'FileDateTime' },
	getattr: { empty: 53, display: 'GetAttr' },
	dir: { blank: 53, display: 'Dir' },
};

const FUNCTIONS: ReadonlySet<string> = new Set(['filelen', 'filedatetime', 'getattr', 'dir']);

const ERROR_TEXT: Readonly<Record<number, string>> = { 53: 'File not found', 75: 'Path/File access error', 76: 'Path not found' };

export function checkEmptyFilePaths(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		let valuesAt: ReturnType<typeof knownLocalLiteralValuesAt> | undefined;
		let consts: ReadonlyMap<string, string> | undefined;
		let sourceNames: ReturnType<typeof sourceNameScopeFor> | undefined;
		return (stmt) => {
			const stringOf = (tok: VbaToken): { value: string; held: boolean } | undefined => {
				if (tok.kind === 'stringLiteral') {
					return { value: stringLiteralValue(tok.rawText), held: false };
				}
				const lower = tokenName(tok)?.toLowerCase();
				if (!lower) {
					return undefined;
				}
				const local = (valuesAt ??= knownLocalLiteralValuesAt(source, member, symbols, activity))(stmt).get(lower);
				if (local) {
					return local.kind === 'string' && !local.contentMutated ? { value: local.value as string, held: true } : undefined;
				}
				const constant = (consts ??= stringConstantsInScope(symbols, member)).get(lower);
				return constant === undefined ? undefined : { value: constant, held: true };
			};
			const check = (spanStart: number, which: string, path: readonly VbaToken[]): void => {
				const errors = ERRORS[which];
				if (path.length !== 1) {
					return;
				}
				const known = stringOf(path[0]);
				if (!known) {
					return;
				}
				const kind = known.value === '' ? 'empty' : /^ +$/.test(known.value) ? 'blank' : undefined;
				const error = kind ? errors[kind] : undefined;
				if (!kind || error === undefined) {
					return;
				}
				const given = known.held ? `'${path[0].rawText}', which holds ${JSON.stringify(known.value)} here` : kind === 'empty' ? 'an empty path' : 'a path of spaces';
				const span: Span = { start: spanStart + path[0].start, end: spanStart + path[0].end };
				push('emptyFilePath', `${errors.display} is given ${given}. This will raise Run-time error '${error}': ${ERROR_TEXT[error]}.`, span);
			};
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
				const head = tokenText(toks[0]);
				for (const path of statementPaths(head, toks)) {
					check(span.start, head, path);
				}
				for (let i = 0; i + 1 < toks.length; i++) {
					const name = tokenText(toks[i]);
					if (!FUNCTIONS.has(name) || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
						continue;
					}
					if (toks[i - 1]?.rawText !== '.' && runtimeCallableSourceShadowed(name, (sourceNames ??= sourceNameScopeFor(symbols, member, projectVisibleSymbols)))) {
						continue;
					}
					const close = matchParenFrom(toks, i + 1);
					const args = close > 0 ? splitTopLevelTokenGroups(toks.slice(i + 2, close), 0, ',') : [];
					if (args.length >= 1 && (name === 'dir' || args.length === 1)) {
						check(span.start, name, args[0]);
					}
				}
			}
		};
	};
}

/** The path arguments a file statement names. */
function statementPaths(head: string, toks: readonly VbaToken[]): VbaToken[][] {
	switch (head) {
		case 'open': {
			const end = toks.findIndex((tok, k) => k > 0 && ['for', 'access', 'shared', 'lock', 'as'].includes(tokenText(tok)));
			return end > 1 ? [toks.slice(1, end)] : [];
		}
		case 'mkdir': case 'chdir': case 'rmdir': case 'kill':
			return toks[1]?.rawText === '=' ? [] : [toks.slice(1)];
		case 'setattr': case 'filecopy': {
			if (toks[1]?.rawText === '=') {
				return [];
			}
			const args = splitTopLevelTokenGroups(toks.slice(1), 0, ',');
			return args.length === 2 ? (head === 'setattr' ? [args[0]] : args) : [];
		}
		case 'name': {
			const as = toks.findIndex((tok) => tokenText(tok) === 'as');
			return as > 1 && toks[1]?.rawText !== '=' ? [toks.slice(1, as)] : [];
		}
		default:
			return [];
	}
}
