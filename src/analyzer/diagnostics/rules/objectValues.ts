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
// A class's default member read the wrong way (issue #256, measured in
// Excel 16.0): one whose first parameter is required, read with no
// argument, raises 449, Argument not optional; one with no parameter that
// returns a Collection gives the Collection, whose own default needs an
// index, 450. A class with no default member indexed, `c(1)`, raises 438.
// For Each over a class asks its -4 member (`VB_UserMemId = -4`) for an
// enumerator: with none it raises 438, and with one returning a
// Collection rather than an object it raises 451.
//
// `Set o = c`, passing `c` to a Variant parameter and `c Is Nothing` read no
// value and run. The write side, `Let c = ...`, is set-required's.

import type { VbaToken } from '../../lexer/tokenKinds';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { BodyNode, ProcedureNode, Span } from '../../parser/nodes';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../../symbols/symbolModel';
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
		const arrays = new Set((procedureSymbolFor(symbols, proc)?.children ?? []).filter((child) => child.isArray).map((child) => child.name.toLowerCase()));
		for (const child of symbols.root.children ?? []) {
			if (child.isArray && !procedureSymbolFor(symbols, proc)?.children?.some((own) => own.name.toLowerCase() === child.name.toLowerCase())) {
				arrays.add(child.name.toLowerCase());
			}
		}
		const classOf = (lower: string): VbaProjectClassMembers | undefined =>
			lower === proc.name.toLowerCase() || arrays.has(lower) ? undefined : projectClass(env.get(lower), memberCtx);
		checkForEachEnumerators(proc.body, classOf, push);
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
				for (const hit of indexedWithoutDefault(statementTokens(source, span), classOf)) {
					push('objectDefaultValue', hit.message, { start: span.start + hit.tok.start, end: span.start + hit.tok.end });
				}
				for (const read of valueReads(source, span, stmt.kind === 'Statement' && stmt.singleLineIfBranches !== undefined && span === stmt.span, isObjectVariable, isTypedValue)) {
					const lower = tokenName(read.tok)!.toLowerCase();
					const cls = classOf(lower);
					const wrongWay = cls && !read.operator && !read.intoTypedValue ? defaultReadProblem(cls) : undefined;
					if (wrongWay) {
						push('objectDefaultValue', `'${read.tok.rawText}' is ${article(cls!.name)} ${cls!.name}, ${wrongWay}`, { start: span.start + read.tok.start, end: span.start + read.tok.end });
						continue;
					}
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

/** A project class's members, when the list is complete. */
function projectClass(type: string | undefined, memberCtx: MemberCompletionContext): VbaProjectClassMembers | undefined {
	const lower = type?.trim().split('.').pop()?.toLowerCase();
	const found = lower ? (memberCtx.projectClassMembers ?? []).find((candidate) => candidate.kind === 'class' && candidate.name.toLowerCase() === lower) : undefined;
	return found?.exhaustive === true ? found : undefined;
}

/** The DISPID a member's attribute gives it: 0 for the default, -4 for the enumerator. */
function dispatchId(member: VbaProjectClassMember): number | undefined {
	const attr = (member.attributes ?? []).find((candidate) => /^vb_(var)?usermemid$/i.test(candidate.name));
	const raw = attr?.valueRaw.trim() ?? '';
	const value = raw.startsWith('-') ? parseVbaIntegerLiteral(raw.slice(1)) : parseVbaIntegerLiteral(raw);
	return value === undefined ? undefined : raw.startsWith('-') ? -value : value;
}

/** Why reading a class's default member with no argument fails, or undefined. */
function defaultReadProblem(cls: VbaProjectClassMembers): string | undefined {
	const member = cls.members.find((candidate) => candidate.defaultMember);
	if (!member) {
		return undefined;
	}
	const first = /^[^(]*\(([^,)]*)/.exec(member.signature ?? '')?.[1]?.trim() ?? '';
	// The signature writes an Optional parameter in brackets: `Item([i As Long = 1])`.
	if (first !== '' && !first.startsWith('[') && !/^paramarray\b/i.test(first)) {
		return `whose default member ${member.name} takes an argument this read does not give. This will raise Run-time error '449': Argument not optional.`;
	}
	if (first === '' && normalizeType(member.returns) === 'collection') {
		return `whose default member ${member.name} returns a Collection, and a Collection's default member Item needs an index. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`;
	}
	return undefined;
}

/** `c(1)` on a class with no default member to take the index. */
function indexedWithoutDefault(
	toks: readonly VbaToken[],
	classOf: (lower: string) => VbaProjectClassMembers | undefined,
): Array<{ tok: VbaToken; message: string }> {
	const out: Array<{ tok: VbaToken; message: string }> = [];
	for (let i = 0; i + 1 < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const cls = lower && toks[i + 1].rawText === '(' && toks[i - 1]?.rawText !== '.' ? classOf(lower) : undefined;
		if (cls && !cls.members.some((member) => member.defaultMember)) {
			out.push({ tok: toks[i], message: `'${toks[i].rawText}' is ${article(cls.name)} ${cls.name}, which has no default member to take an index. This will raise Run-time error '438': Object doesn't support this property or method.` });
		}
	}
	return out;
}

/** `For Each v In c` over a class: the -4 member it needs, and what that returns. */
function checkForEachEnumerators(
	body: readonly BodyNode[],
	classOf: (lower: string) => VbaProjectClassMembers | undefined,
	push: PushFn,
): void {
	for (const node of body) {
		if (node.kind === 'ForBlock' && node.each && node.sourceExpressionSpan) {
			const over = node.sourceExpression?.trim() ?? '';
			const cls = /^[\p{L}_][\p{L}\p{N}_]*$/u.test(over) ? classOf(over.toLowerCase()) : undefined;
			const enumerator = cls?.members.find((member) => dispatchId(member) === -4);
			const returns = normalizeType(enumerator?.returns);
			if (cls && !enumerator) {
				push('objectDefaultValue', `'${over}' is ${article(cls.name)} ${cls.name}, which has no member marked VB_UserMemId = -4 for For Each to ask for its elements. This will raise Run-time error '438': Object doesn't support this property or method.`, node.sourceExpressionSpan);
			} else if (cls && enumerator && returns === 'collection') {
				push('objectDefaultValue', `'${over}' is ${article(cls.name)} ${cls.name}, whose enumerator ${enumerator.name} returns a Collection, not the enumerator object For Each needs. This will raise Run-time error '451': Property let procedure not defined and property get procedure did not return an object.`, node.sourceExpressionSpan);
			}
		}
		if ('body' in node && Array.isArray(node.body)) {
			checkForEachEnumerators(node.body as BodyNode[], classOf, push);
		}
	}
}

function article(type: string): string {
	return /^[aeiou]/i.test(type) ? 'an' : 'a';
}
