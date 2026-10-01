// Rule: an object read as a value when its type has no default member to
// give one (issue #183). Measured in Excel 16.0 (build 20326, 2026-09-29);
// each compiles and raises every time it runs.
//
//  - A class with no default member: `s = c`, `v = c` into a Variant,
//    `Debug.Print c`, `c & "x"`, `c + 1`, `If c = 1` -> 438, Object doesn't
//    support this property or method. So do a Worksheet and a Workbook.
//  - A Collection, whose default member Item needs an index: `v = c` and
//    `Debug.Print c` -> 450, Wrong number of arguments or invalid property
//    assignment. With an operator, or into a String or other typed value,
//    it does not compile, which is collection-operand's. Excel's collections
//    whose default is their Item do the same: Hyperlinks, Areas, Borders,
//    Windows, Workbooks, Shapes (issue #221).
//  - Excel's other objects with no default member raise 438 as a Worksheet
//    does: Workbook, Font, Interior, Validation, Window, PageSetup, Border,
//    Shape, Hyperlink (issue #221).
//  - An object variable still Nothing raises 91 first.
//
// `Set o = c`, passing `c` to a Variant parameter and `c Is Nothing` read no
// value and run. The write side, `Let c = ...`, is set-required's.

import type { VbaToken } from '../../lexer/tokenKinds';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ProcedureNode, Span } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { isKnownObjectAssignmentType, isKnownScalarType, normalizeType, objectLetAssignmentVerdict, objectValueNeedsIndex, typeEnvironmentFor } from '../typeInference';
import {
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

export function checkObjectDefaultValues(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleAutoInstanced = new Set<string>();
	const moduleNames = new Set((symbols.root.children ?? []).map((child) => child.name.toLowerCase()));
	for (const child of symbols.root.children ?? []) {
		if (child.isAutoInstantiated) {
			moduleAutoInstanced.add(child.name.toLowerCase());
		}
	}
	return (proc: ProcedureNode) => {
		const env = typeEnvironmentFor(symbols, proc);
		const verdicts = new Map<string, ReturnType<typeof objectLetAssignmentVerdict>>();
		const verdictFor = (lower: string): ReturnType<typeof objectLetAssignmentVerdict> => {
			let verdict = verdicts.get(lower);
			if (verdict === undefined) {
				// The procedure's own name is its return value only as a target;
				// read, it is a recursive call.
				const type = lower === proc.name.toLowerCase() ? undefined : env.get(lower);
				verdict = type ? objectLetAssignmentVerdict(type, memberCtx) : 'unknown';
				verdicts.set(lower, verdict);
			}
			return verdict;
		};
		// An `As New` variable is never Nothing when read. A local shadows a
		// module-level variable of the same name.
		const autoInstanced = new Set(moduleAutoInstanced);
		for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
			if (child.isAutoInstantiated) {
				autoInstanced.add(child.name.toLowerCase());
			} else {
				autoInstanced.delete(child.name.toLowerCase());
			}
		}
		const isObjectVariable = (name: string): boolean => {
			const type = env.get(name.toLowerCase());
			return type !== undefined && isKnownObjectAssignmentType(type, memberCtx);
		};
		// A Let target of a declared scalar type: `s = c` with s a String.
		const isTypedValue = (name: string): boolean => {
			const lower = name.toLowerCase();
			const type = normalizeType(lower === proc.name.toLowerCase() ? proc.returnType : env.get(lower));
			return type !== undefined && isKnownScalarType(type);
		};
		const isCollection = (lower: string): boolean => lower !== proc.name.toLowerCase() && normalizeType(env.get(lower)) === 'collection';
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				for (const hit of collectionArguments(statementTokens(source, span), isCollection, moduleNames)) {
					const at = { start: span.start + hit.start, end: span.start + hit.end };
					if (hit.compiles) {
						push('objectDefaultValue', `${hit.what} is a Collection: its default member Item needs an index, so ${hit.fn} has no value to read. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`, at);
					} else {
						push('collectionOperand', `${hit.what} is a Collection: its default member Item needs an index, so ${hit.fn} has no value to take. This is a VBE compile error: Argument not optional.`, at);
					}
				}
				const created = newObjectLetIntoVariant(source, span, env, proc);
				if (created) {
					const verdict = objectLetAssignmentVerdict(created.type, memberCtx);
					if (verdict === 'noDefault' || (verdict === 'argument' && normalizeType(created.type) === 'collection')) {
						push(
							'objectDefaultValue',
							verdict === 'noDefault'
								? `'New ${created.type}' is assigned without Set, so its value is read, and ${created.type} has no default member to give one. This will raise Run-time error '438': Object doesn't support this property or method.`
								: `'New ${created.type}' is assigned without Set, so its value is read, and a Collection's default member Item needs an index. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`,
							created.span,
						);
					}
				}
				for (const read of valueReads(source, span, stmt.kind === 'Statement' && stmt.singleLineIfBranches !== undefined && span === stmt.span, isObjectVariable, isTypedValue)) {
					const lower = tokenName(read.tok)!.toLowerCase();
					const verdict = verdictFor(lower);
					if (verdict !== 'noDefault' && verdict !== 'argument') {
						continue;
					}
					const type = env.get(lower)!;
					// A default member that needs an index raises 450 read as a
					// value; with an operator, or into a typed value, it is a
					// compile error, collection-operand's.
					if (verdict === 'argument' && (read.operator || read.intoTypedValue || !objectValueNeedsIndex(type, memberCtx))) {
						continue;
					}
					const nothing = autoInstanced.has(lower) ? '' : `, or '91' while it is Nothing`;
					const at: Span = { start: span.start + read.tok.start, end: span.start + read.tok.end };
					push(
						'objectDefaultValue',
						verdict === 'noDefault'
							? `'${read.tok.rawText}' is ${article(type)} ${type}, which has no default member, so it has no value to read here. This will raise Run-time error '438': Object doesn't support this property or method${nothing}.`
							: `'${read.tok.rawText}' is ${article(type)} ${type}: its default member Item needs an index, so it has no value to read here. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment${nothing}.`,
						at,
					);
				}
			}
		};
	};
}

/**
 * Built-ins whose argument takes a value, measured in Excel 16.0 with a
 * Collection (issue #242): a typed parameter refuses it while compiling,
 * `Len(c)`, `CStr(c)`, `Abs(c)`; a Variant one asks the Item for a value
 * at run time and raises 450, `InStr(c, "a")`, `Format(c)`, `Hex(c)`.
 * `TypeName(c)` and `IsNumeric(c)` read no value.
 */
const REFUSING_BUILTINS: ReadonlySet<string> = new Set([
	'len', 'cstr', 'val', 'clng', 'cdbl', 'cint', 'cbool', 'cdate', 'trim$', 'ucase$', 'lcase$',
	'instrrev', 'asc', 'chr', 'abs',
]);
const VALUE_READING_BUILTINS: ReadonlySet<string> = new Set([
	'instr', 'format', 'ucase', 'lcase', 'trim', 'ltrim', 'rtrim', 'left', 'right', 'mid', 'cvar', 'strcomp', 'hex',
]);

/**
 * A Collection, a variable or `New Collection`, as the first argument of
 * one of those built-ins. Offsets are the statement's.
 */
function collectionArguments(
	toks: readonly VbaToken[],
	isCollection: (lower: string) => boolean,
	moduleNames: ReadonlySet<string>,
): Array<{ start: number; end: number; fn: string; what: string; compiles: boolean }> {
	const out: Array<{ start: number; end: number; fn: string; what: string; compiles: boolean }> = [];
	for (let i = 0; i + 2 < toks.length; i++) {
		// `Trim$(` lexes as Trim and a `$` of its own.
		const suffixed = toks[i + 1].rawText === '$';
		const fn = toks[i].rawText.toLowerCase() + (suffixed ? '$' : '');
		const open = suffixed ? i + 2 : i + 1;
		const refuses = REFUSING_BUILTINS.has(fn);
		if ((!refuses && !VALUE_READING_BUILTINS.has(fn)) || toks[open]?.rawText !== '(' || moduleNames.has(toks[i].rawText.toLowerCase())) {
			continue;
		}
		const qualified = toks[i - 1]?.rawText === '.';
		if (qualified && tokenText(toks[i - 2]) !== 'vba') {
			continue;
		}
		const a = toks[open + 1];
		if (!a) {
			continue;
		}
		const created = tokenText(a) === 'new' && tokenText(toks[open + 2]) === 'collection';
		const last = created ? open + 2 : open + 1;
		const closes = toks[last + 1]?.rawText === ')' || toks[last + 1]?.rawText === ',';
		const name = tokenName(a)?.toLowerCase();
		if (!closes || (!created && (!name || !isCollection(name)))) {
			continue;
		}
		out.push({
			start: a.start,
			end: toks[last].end,
			fn: toks[i].rawText + (suffixed ? '$' : ''),
			what: created ? "'New Collection'" : `'${a.rawText}'`,
			compiles: !refuses,
		});
	}
	return out;
}

/**
 * The names a statement reads as a value: the whole value of a Let
 * (`s = c`), an item Debug.Print prints (`Debug.Print "a"; c`), and an operand
 * of a scalar operator (`c & "x"`). A name followed by `(` or `.`, or after a
 * `.`, is a call or a member access, not the object's own value.
 */
function valueReads(
	source: string,
	span: Span,
	ifHead: boolean,
	isObjectVariable: (name: string) => boolean,
	isTypedValue: (name: string) => boolean,
): Array<{ tok: VbaToken; operator: boolean; intoTypedValue?: boolean }> {
	const toks = statementTokens(source, span);
	const first = firstExecutableTokenIndex(toks);
	const head = tokenText(toks[first]);
	if (head === 'set') {
		return [];
	}
	const out: Array<{ tok: VbaToken; operator: boolean; intoTypedValue?: boolean }> = [];
	const plainName = (i: number): boolean => tokenName(toks[i]) !== undefined
		&& toks[i - 1]?.rawText !== '.' && toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.';
	const target = bareAssignmentTarget(source, span);
	const eq = target ? toks.findIndex((tok) => tok.rawText === '=') : -1;
	if (target) {
		const value = target.valueTokens.filter((tok) => tok.kind !== 'comment');
		// A Let into an object variable is set-required's.
		if (value.length === 1 && eq + 1 === toks.indexOf(value[0]) && plainName(eq + 1) && !isObjectVariable(target.name)) {
			out.push({ tok: value[0], operator: false, intoTypedValue: isTypedValue(target.name) });
		}
	}
	if (head === 'debug' && toks[first + 1]?.rawText === '.' && tokenText(toks[first + 2]) === 'print') {
		let depth = 0;
		for (let i = first + 3; i < toks.length; i++) {
			const raw = toks[i].rawText;
			depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
			const before = toks[i - 1]?.rawText;
			const after = toks[i + 1]?.rawText;
			const alone = (i === first + 3 || before === ',' || before === ';') && (after === undefined || after === ',' || after === ';' || toks[i + 1].kind === 'comment');
			if (depth === 0 && alone && plainName(i)) {
				out.push({ tok: toks[i], operator: false });
			}
		}
	}
	const then = ifHead ? toks.findIndex((tok) => tokenText(tok) === 'then') : -1;
	const limit = then > 0 ? then : toks.length;
	for (let i = first; i < limit; i++) {
		if (i === eq - 1 || !plainName(i) || out.some((read) => read.tok === toks[i])) {
			continue;
		}
		const previous = i - 1 === eq ? undefined : toks[i - 1];
		const isOperator = (tok: VbaToken | undefined): boolean => tok !== undefined
			&& ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod');
		if (isOperator(toks[i + 1]) || isOperator(previous)) {
			out.push({ tok: toks[i], operator: true });
		}
	}
	return out;
}

/**
 * `v = New Collection` into a Variant, or into the Function's own result: a
 * Let, which reads the new object's default value (issue #219, measured in
 * Excel 16.0; `Set v = New Collection` runs). Into a typed scalar the VBE
 * refuses it while compiling, which is not judged here.
 */
function newObjectLetIntoVariant(
	source: string,
	span: Span,
	env: ReadonlyMap<string, string>,
	proc: ProcedureNode,
): { type: string; span: Span } | undefined {
	const target = bareAssignmentTarget(source, span);
	if (!target) {
		return undefined;
	}
	const value = target.valueTokens.filter((tok) => tok.kind !== 'comment');
	if (value.length !== 2 || tokenText(value[0]) !== 'new' || !tokenName(value[1])) {
		return undefined;
	}
	const lower = target.name.toLowerCase();
	const isResult = proc.procKind === 'Function' && lower === proc.name.toLowerCase();
	if (!isResult && !env.has(lower)) {
		return undefined;
	}
	const declared = normalizeType(isResult ? proc.returnType : env.get(lower));
	if ((declared !== undefined && declared !== 'variant') || (isResult && proc.typeSuffix)) {
		return undefined;
	}
	return { type: value[1].rawText, span: { start: span.start + value[0].start, end: span.start + value[1].end } };
}

function article(type: string): string {
	return /^[aeiou]/i.test(type) ? 'an' : 'a';
}
