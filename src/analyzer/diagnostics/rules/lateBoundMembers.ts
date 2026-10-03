// Rule: a member the VBE binds at run time, on an object whose class the code
// makes plain and whose member list is complete (issue #121). Measured in
// Excel 16.0 (build 20326, 2026-09-26); each compiles and raises 438, "Object
// doesn't support this property or method", every time it runs.
//
//  - `Application.Zzq`: Application is extensible, so the VBE compiles any
//    name on it (worksheet functions such as Application.Match are ordinary
//    VBA there), and a name that is neither an Application member nor a
//    WorksheetFunction raises when it runs. Excel only: its model lists every
//    member, hidden ones included.
//  - `Dim o As Object: Set o = New Collection: o.Foo`: a late-bound variable
//    holding a class with a known member list. Collection has Add, Count,
//    Item and Remove; a project class module has its public members.
//
// Issue #224 (measured in Excel 16.0): the class also reaches the variable
// from one declared as it, `Set o = c`, where c raises 91 instead while it is
// Nothing. A Private member is not on the list (438). A property with a Get
// and no Let raises 451 when assigned, and one with a Let and no Get 450 when
// read.

import { getExcelObjectModel, type HostObjectModel } from '../../host/excelObjectModel';
import { getHostMembers, getHostType } from '../../host/hostModel';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { projectTypeAt, resolveReceiverTypeAt } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { jumpTargetLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ForBlockNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { HELD_VALUE, heldObjectsAt } from '../heldObjects';
import { isLeafStatement } from '../../parser/nodes';
import { walkEnteringBlocks } from '../dataflow';
import { bodyMayLeaveLoop, namesIn } from './shared';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { buildModuleTypeSignatures, inferExpressionType, isKnownScalarType, normalizeType, objectAssignmentIncompatibilityReason, sourceNameScopeFor, stringLiteralValue, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	forEachStatement,
	rawExpressionTokens,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { matchParenFrom, splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import { vbscriptPatternError } from './lateBoundObjects';

const COLLECTION_MEMBERS: ReadonlySet<string> = new Set(['add', 'count', 'item', 'remove']);

/** VBScript's RegExp, as CreateObject("VBScript.RegExp") gives it (issue #477). */
const REGEXP_CLASS = {
	display: 'RegExp',
	members: new Set(['pattern', 'global', 'ignorecase', 'multiline', 'test', 'execute', 'replace']),
};

/**
 * ProgIDs everyday macros create, lowercased. A ProgID one letter away from
 * one of these, with the same parts, is taken as a misspelling of it, and
 * the bare last part ("Dictionary") is no ProgID at all (issue #477,
 * measured in Excel 16.0: each raises 429).
 */
const KNOWN_PROGIDS: readonly string[] = [
	'scripting.dictionary', 'scripting.filesystemobject', 'vbscript.regexp', 'wscript.shell', 'wscript.network', 'shell.application',
	'adodb.connection', 'adodb.recordset', 'adodb.command', 'adodb.stream',
	'msxml2.domdocument', 'msxml2.domdocument.3.0', 'msxml2.domdocument.4.0', 'msxml2.domdocument.5.0', 'msxml2.domdocument.6.0',
	'msxml2.xmlhttp', 'msxml2.xmlhttp.3.0', 'msxml2.xmlhttp.6.0', 'msxml2.serverxmlhttp', 'msxml2.serverxmlhttp.6.0', 'winhttp.winhttprequest.5.1',
	'excel.application', 'word.application', 'powerpoint.application', 'outlook.application', 'access.application',
];

/** Why a ProgID literal names nothing, or undefined when it may be registered. */
function progIdProblem(progId: string): string | undefined {
	const lower = progId.trim().toLowerCase();
	if (lower === '') {
		return 'an empty ProgID names no class';
	}
	if (KNOWN_PROGIDS.includes(lower)) {
		return undefined;
	}
	const bare = KNOWN_PROGIDS.find((known) => !lower.includes('.') && known.split('.')[1] === lower);
	if (bare) {
		return `"${progId}" lacks its library: the ProgID is "${bare}"`;
	}
	const near = KNOWN_PROGIDS.find((known) => known.split('.').length === lower.split('.').length && oneLetterApart(known, lower));
	return near ? `"${progId}" is one letter away from "${near}", and no class has that ProgID` : undefined;
}

/** Whether two strings differ by one inserted, deleted or changed letter. */
function oneLetterApart(a: string, b: string): boolean {
	if (Math.abs(a.length - b.length) > 1 || a === b) {
		return false;
	}
	let i = 0;
	while (i < a.length && i < b.length && a[i] === b[i]) {
		i++;
	}
	const letter = (c: string | undefined): boolean => c === undefined || /[a-z]/.test(c);
	if (a.length === b.length) {
		return letter(a[i]) && letter(b[i]) && a.slice(i + 1) === b.slice(i + 1);
	}
	const [long, short] = a.length > b.length ? [a, b] : [b, a];
	return letter(long[i]) && long.slice(i + 1) === short.slice(i);
}

/**
 * What VBScript's RegExp refuses in a pattern, with the error it raises
 * (issue #477, measured in Excel 16.0): an unclosed group (5020), an
 * unclosed class (5019), a quantifier with nothing to repeat (5018), and
 * any other fault, a lookbehind or a named group among them (5017).
 */
function regExpPatternProblem(pattern: string): { error: string; text: string } | undefined {
	switch (vbscriptPatternError(pattern)) {
		case 5020: return { error: "'5020': Application-defined or object-defined error (VBScript: Expected ')' in regular expression)", text: 'an unclosed group' };
		case 5019: return { error: "'5019': Application-defined or object-defined error (VBScript: Expected ']' in regular expression)", text: 'an unclosed character class' };
		case 5018: return { error: "'5018': Application-defined or object-defined error (VBScript: Unexpected quantifier)", text: 'a quantifier with nothing to repeat' };
		case 5017: return { error: "'5017': Application-defined or object-defined error (VBScript: Syntax error in regular expression)", text: 'a form VBScript does not read, such as a lookbehind' };
		default: return undefined;
	}
}

/** Names a late-bound local is known to hold: the class display name and its members. */
interface KnownClass {
	display: string;
	members: ReadonlySet<string>;
	/** Properties with a Get and no Let or Set: assigning one raises 451. */
	readOnly?: ReadonlySet<string>;
	/** Properties with a Let and no Get: reading one raises 450. */
	writeOnly?: ReadonlySet<string>;
	/** Properties with only a Set: reading one raises 450 too (issue #414). */
	setOnly?: ReadonlySet<string>;
	/** Properties with a Get and a Set and no Let: a Let of one raises 438 (issue #685). */
	noLet?: ReadonlySet<string>;
	/** Subs: assigning one raises 450, reading one with arguments 451 (issue #414). */
	subs?: ReadonlySet<string>;
	/** Fields of a scalar type, by lowercased name: a member of one raises 424 (issue #414). */
	scalarFields?: ReadonlyMap<string, string>;
	/** Set from a variable that may still be Nothing: 91 before 438. */
	mayBeNothing?: boolean;
	/** The parameters of each method, by lowercased name, where they are known (issue #485). */
	params?: ReadonlyMap<string, readonly KnownParam[]>;
	/** A RegExp's pattern, where the code set it to a literal (issue #477). */
	pattern?: string;
}

interface KnownParam {
	name: string;
	optional: boolean;
	paramArray: boolean;
}

/** A Collection's methods as its type library declares them. */
const COLLECTION_PARAMS: ReadonlyMap<string, readonly KnownParam[]> = new Map([
	['add', [
		{ name: 'Item', optional: false, paramArray: false },
		{ name: 'Key', optional: true, paramArray: false },
		{ name: 'Before', optional: true, paramArray: false },
		{ name: 'After', optional: true, paramArray: false },
	]],
	['item', [{ name: 'Index', optional: false, paramArray: false }]],
	['remove', [{ name: 'Index', optional: false, paramArray: false }]],
	['count', []],
]);

/**
 * What a call of a known member with these arguments raises, or undefined:
 * a named argument it has no parameter for (448), more arguments than it
 * takes (450), a required one missing (449), and an argument to a Count
 * that takes none (451). The call is `o.M(...)`, or `o.M ...` as the
 * statement. `toks[at]` is the receiver.
 */
function argumentRefusal(toks: readonly VbaToken[], at: number, params: readonly KnownParam[], memberName: string, property: boolean): string | undefined {
	const open = at + 3;
	let args: VbaToken[][] | undefined;
	if (toks[open]?.rawText === '(') {
		const close = matchParenFrom(toks, open);
		if (close < 0) {
			return undefined;
		}
		args = close === open + 1 ? [] : splitTopLevelTokenGroups([...toks], open + 1, ',', close);
	} else if (at === 0 && toks.length > open && toks[open].rawText !== '=' && toks[open].rawText !== '.') {
		args = splitTopLevelTokenGroups([...toks].filter((tok) => tok.kind !== 'comment'), open, ',', toks.filter((tok) => tok.kind !== 'comment').length);
	} else if (at === 0 && toks.length === open) {
		args = [];
	} else if (at > 0 && toks[open]?.rawText !== '=' && toks[open]?.rawText !== '.' && toks[open]?.rawText !== '!') {
		// Read with no parentheses, `x = o.Idx`: no arguments (issue #685).
		args = [];
	}
	if (!args) {
		return undefined;
	}
	if (property) {
		return args.length > 0 ? `its ${memberName} takes no argument. This will raise Run-time error '451': Property let procedure not defined and property get procedure did not return an object` : undefined;
	}
	const named = args.filter((arg) => arg.length >= 3 && arg[1].rawText === ':=');
	for (const arg of named) {
		if (!params.some((param) => param.name.toLowerCase() === arg[0].rawText.toLowerCase())) {
			return `its ${memberName} has no parameter named '${arg[0].rawText}'. This will raise Run-time error '448': Named argument not found`;
		}
	}
	const positional = args.length - named.length;
	if (named.length === 0 && !params.some((param) => param.paramArray) && positional > params.length) {
		return `its ${memberName} takes at most ${params.length} argument(s), and ${positional} are passed. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment`;
	}
	const given = new Set(named.map((arg) => arg[0].rawText.toLowerCase()));
	const missing = params.find((param, k) => !param.optional && !param.paramArray && !given.has(param.name.toLowerCase())
		&& (k >= positional || (args![k] !== undefined && args![k].filter((tok) => tok.kind !== 'comment').length === 0)));
	return missing ? `its ${memberName} needs '${missing.name}', which is not passed. This will raise Run-time error '449': Argument not optional` : undefined;
}

/** The parameters a member signature lists: `M(ByVal a As Long, [ByVal b As Long])`. */
function signatureParams(signature: string): KnownParam[] | undefined {
	const open = signature.indexOf('(');
	let depth = 0;
	let close = -1;
	for (let i = open; open >= 0 && i < signature.length; i++) {
		depth += signature[i] === '(' ? 1 : signature[i] === ')' ? -1 : 0;
		if (depth === 0) {
			close = i;
			break;
		}
	}
	if (close < 0) {
		return undefined;
	}
	const list = signature.slice(open + 1, close).trim();
	const params: KnownParam[] = [];
	for (const raw of list === '' ? [] : list.split(',')) {
		const text = raw.trim();
		const optional = text.startsWith('[') || /^optional\b/i.test(text);
		const words = text.replace(/[[\]]/g, '').split(/\s+/).filter((word) => !/^(optional|byval|byref|paramarray)$/i.test(word));
		const name = /^[A-Za-z_][A-Za-z0-9_]*/.exec(words[0] ?? '')?.[0];
		if (!name) {
			return undefined;
		}
		params.push({ name, optional, paramArray: /\bparamarray\b/i.test(text) });
	}
	return params;
}

/** The Scripting.FileSystemObject's members, as its type library lists them. */
const FSO_CLASS = {
	display: 'FileSystemObject',
	members: new Set([
		'drives', 'buildpath', 'copyfile', 'copyfolder', 'createfolder', 'createtextfile', 'deletefile', 'deletefolder', 'driveexists',
		'fileexists', 'folderexists', 'getabsolutepathname', 'getbasename', 'getdrive', 'getdrivename', 'getextensionname', 'getfile',
		'getfilename', 'getfileversion', 'getfolder', 'getparentfoldername', 'getspecialfolder', 'getstandardstream', 'gettempname',
		'movefile', 'movefolder', 'opentextfile',
	]),
};

/** The class a CreateObject of a ProgID literal gives, where its members are known. */
function progIdClass(progId: string): KnownClass | undefined {
	const lower = progId.trim().toLowerCase();
	return lower === 'vbscript.regexp' ? { ...REGEXP_CLASS } : lower === 'scripting.filesystemobject' ? { ...FSO_CLASS } : undefined;
}

/**
 * Faults a literal shows on these objects (issue #477, measured in Excel
 * 16.0): CreateObject and GetObject of a ProgID no class has (429), a
 * RegExp pattern VBScript refuses at Test, Execute or Replace, Null given to
 * them (13), Global, IgnoreCase or MultiLine set to text that is no Boolean
 * (13), and an IOMode OpenTextFile does not take (5).
 */
function checkProgIdObjects(base: number, toks: readonly VbaToken[], held: ReadonlyMap<string, KnownClass>, push: PushFn): void {
	const at = (tok: VbaToken): Span => ({ start: base + tok.start, end: base + tok.end });
	for (let i = 0; i + 2 < toks.length; i++) {
		const word = tokenText(toks[i]);
		if ((word === 'createobject' || word === 'getobject') && toks[i + 1].rawText === '(' && toks[i - 1]?.rawText !== '.') {
			const close = matchParenFrom([...toks], i + 1);
			const args = close > i + 1 ? splitTopLevelTokenGroups([...toks], i + 2, ',', close) : [];
			const arg = word === 'createobject' ? args[0] : args[1];
			const literal = arg?.length === 1 && arg[0].kind === 'stringLiteral' ? arg[0] : undefined;
			const problem = literal ? progIdProblem(stringLiteralValue(literal.rawText)) : undefined;
			if (literal && problem) {
				push('runtimeArgumentValue', `${word === 'createobject' ? 'CreateObject' : 'GetObject'}: ${problem}. This will raise Run-time error '429': ActiveX component can't create object.`, at(literal));
			}
			continue;
		}
		const known = held.get(tokenName(toks[i])?.toLowerCase() ?? '');
		if (!known || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.') {
			continue;
		}
		const member = tokenText(toks[i + 2]);
		if (known.display === 'RegExp') {
			if (['test', 'execute', 'replace'].includes(member) && toks[i + 3]?.rawText === '(') {
				const problem = known.pattern !== undefined ? regExpPatternProblem(known.pattern) : undefined;
				const close = matchParenFrom([...toks], i + 3);
				const first = close > i + 4 ? splitTopLevelTokenGroups([...toks], i + 4, ',', close)[0] : undefined;
				if (problem) {
					push('runtimeArgumentValue', `The pattern "${known.pattern}" has ${problem.text}. This will raise Run-time error ${problem.error}.`, at(toks[i + 2]));
				} else if (first?.length === 1 && tokenText(first[0]) === 'null') {
					push('runtimeArgumentValue', `RegExp.${toks[i + 2].rawText} takes a String, and Null is none. This will raise Run-time error '13': Type mismatch.`, at(first[0]));
				}
			} else if (['global', 'ignorecase', 'multiline'].includes(member) && i === 0 && toks[3]?.rawText === '=' && toks[4]?.kind === 'stringLiteral' && toks.length === 5) {
				const text = stringLiteralValue(toks[4].rawText);
				if (!/^\s*(true|false)\s*$/i.test(text) && !/\d/.test(text)) {
					push('runtimeArgumentValue', `RegExp.${toks[2].rawText} takes True or False, and "${text}" is neither. This will raise Run-time error '13': Type mismatch.`, at(toks[4]));
				}
			}
		}
	}
}

export function checkRuntimeMemberNotFound(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const model = memberCtx.model;
	const applicationSurface = excelApplicationSurface(model);
	const rangeSurface = applicationSurface ? excelRangeSurface(model) : undefined;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const added = controlsAddedIn(source, member.body, activity);
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				checkFormControlNames(source, span.start, toks, memberCtx, added, push);
				checkOpenTypeMembers(source, span.start, toks, env, applicationSurface, rangeSurface, memberCtx, push);
			}
		}, activity);
		checkCollectionItems(source, member, symbols, env, memberCtx, activity, push);
		const autoInstanced = new Set<string>();
		for (const child of procedureSymbolFor(symbols, member)?.children ?? []) {
			if (child.isAutoInstantiated) {
				autoInstanced.add(child.name.toLowerCase());
			}
		}
		// Asked only for the target of a Set: walking the whole environment
		// for every procedure was 5% of a large module's pass (issue #139).
		const isLateBound = (lower: string): boolean => {
			if (!env.has(lower)) {
				return false;
			}
			const normalized = normalizeType(env.get(lower));
			return normalized === 'object' || normalized === 'variant' || normalized === undefined;
		};
		const held = new Map<string, KnownClass>();
		// Blocks are entered with the state they start with (issue #237).
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return; // a Dim inside the body declares, and runs nothing
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (jumpTargetLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				held.clear();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forgetMentioned(toks, held);
				return;
			}
			checkProgIdObjects(node.span.start, toks, held, push);
			checkStatement(source, node.span.start, toks, held, applicationSurface, memberCtx, push);
			// `re.Pattern = "(a"`: the pattern a later Test or Execute reads.
			const target = tokenName(toks[0])?.toLowerCase();
			const regExp = target ? held.get(target) : undefined;
			if (regExp?.display === 'RegExp' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'pattern' && toks[3]?.rawText === '=') {
				const value = toks.slice(4).filter((tok) => tok.kind !== 'comment');
				held.set(target!, { ...regExp, pattern: value.length === 1 && value[0].kind === 'stringLiteral' ? stringLiteralValue(value[0].rawText) : undefined });
				return;
			}
			const set = setAssignmentTarget(source, node.span);
			if (set && isLateBound(set.name.toLowerCase())) {
				const lower = set.name.toLowerCase();
				const value = toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1);
				const source1 = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
				const fromVariable = source1 !== undefined && !isLateBound(source1) ? knownClassNamed(env.get(source1), memberCtx) : undefined;
				const created = value.length === 4 && tokenText(value[0]) === 'createobject' && value[1].rawText === '(' && value[2].kind === 'stringLiteral' && value[3].rawText === ')'
					? progIdClass(stringLiteralValue(value[2].rawText))
					: undefined;
				const known = created ?? (value.length === 2 && tokenText(value[0]) === 'new'
					? knownClassNamed(tokenName(value[1]), memberCtx)
					: fromVariable && { ...fromVariable, mayBeNothing: !autoInstanced.has(source1!) });
				if (known) {
					held.set(lower, known);
				} else {
					held.delete(lower);
				}
				return;
			}
			forgetOtherUses(toks, held);
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map(held),
			restore: (saved) => {
				held.clear();
				for (const [lower, known] of saved) {
					held.set(lower, known);
				}
			},
			forget: (names) => {
				for (const lower of names) {
					held.delete(lower);
				}
			},
			touches: (stmt) => namesIn(source, stmt.span),
		});
	}
}

/**
 * The items a Collection local holds, by class (issue #246, measured in Excel
 * 16.0): `c.Add New Flat1` then `c(1).Radius()` asks a Flat1 for a member it
 * lacks, 438; and `For Each x In c` with x a Round1 Sets each item into x,
 * raising 13 at an item of another class.
 */
function checkCollectionItems(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	env: ReadonlyMap<string, string>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const heldAt = heldObjectsAt(source, proc, symbols, activity);
	forEachStatement(proc.body, (stmt) => {
		const items = heldAt(stmt).items;
		if (items.size === 0 || (stmt.kind === 'Statement' && stmt.singleLineIfBranches)) {
			return;
		}
		const toks = statementTokens(source, stmt.span);
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			const held = lower ? items.get(lower) : undefined;
			if (!held || held.length === 0 || toks[i - 1]?.rawText === '.') {
				continue;
			}
			// `c(1).Member` or `c.Item(1).Member`.
			const open = toks[i + 1]?.rawText === '(' ? i + 1 : toks[i + 1]?.rawText === '.' && tokenText(toks[i + 2]) === 'item' && toks[i + 3]?.rawText === '(' ? i + 3 : -1;
			const close = open >= 0 ? matchParenFrom(toks, open) : -1;
			if (close < 0 || toks[close + 1]?.rawText !== '.' || !tokenName(toks[close + 2])) {
				continue;
			}
			const index = close === open + 2 && toks[open + 1].kind === 'integerLiteral' ? Number(toks[open + 1].rawText) : undefined;
			const className = index !== undefined && index >= 1 && index <= held.length ? held[index - 1]
				: held.every((name) => name.toLowerCase() === held[0].toLowerCase()) ? held[0] : undefined;
			const known = knownClassNamed(className, memberCtx);
			const memberName = tokenName(toks[close + 2])!;
			if (!known || known.members.has(memberName.toLowerCase())) {
				continue;
			}
			const shown = toks.slice(i, close + 1).map((tok) => tok.rawText).join('');
			push('runtimeMemberNotFound', `'${shown}' holds a ${known.display} here, which has no member '${memberName}'. This will raise Run-time error '438': Object doesn't support this property or method.`, { start: stmt.span.start + toks[i].start, end: stmt.span.start + toks[close].end });
		}
	}, activity);
	const signatures = buildModuleTypeSignatures(symbols);
	const sourceNames = sourceNameScopeFor(symbols, proc);
	forEachLoopIn(proc.body, activity, (loop) => {
		const source1 = loop.sourceExpression?.trim().toLowerCase();
		const control = loop.controlVariable?.toLowerCase();
		const held = source1 ? heldAt(loop).items.get(source1) : undefined;
		const expected = control ? env.get(control) : undefined;
		// `For Each c In Worksheets` with c As Range: each sheet is Set into c,
		// and a sheet is no Range (issue #447, measured in Excel 16.0).
		const element = expected && loop.sourceExpressionSpan && !held
			? hostElementType(source, loop.sourceExpressionSpan, env, signatures, sourceNames, memberCtx)
			: undefined;
		if (element && loop.sourceExpressionSpan) {
			const bare = element.replace(/^\w+\./, '');
			const label = `the items of '${loop.sourceExpression!.trim()}', each ${/^[AEIOU]/.test(bare) ? 'an' : 'a'} ${bare}`;
			if (objectAssignmentIncompatibilityReason(expected!, { type: element, label, span: loop.sourceExpressionSpan }, { ...memberCtx, model: memberCtx.model ?? getExcelObjectModel() })) {
				push('assignmentObjectTypeMismatch', `For Each Sets ${label}, into '${loop.controlVariable}', a ${expected}. This will raise Run-time error '13': Type mismatch.`, loop.sourceExpressionSpan);
				return;
			}
		}
		if (!held || !expected || !loop.sourceExpressionSpan) {
			return;
		}
		// A body that may leave the loop may stop before any later item: only
		// the first is certainly Set (issue #356, measured in Excel 16.0).
		const reached = bodyMayLeaveLoop(source, loop.body) ? held.slice(0, 1) : held;
		// A number or string Set into an object variable is Object required, 424
		// (issue #447, measured in Excel 16.0); a Variant takes it.
		const objectControl = normalizeType(expected) !== 'variant' && !isKnownScalarType(normalizeType(expected) ?? '');
		const position = reached.findIndex((name) => (name === HELD_VALUE
			? objectControl
			: objectAssignmentIncompatibilityReason(expected, { type: name, label: name, span: loop.sourceExpressionSpan! }, memberCtx) !== undefined));
		if (position >= 0 && held[position] === HELD_VALUE) {
			push('assignmentObjectTypeMismatch', `For Each Sets each item of '${loop.sourceExpression!.trim()}' into '${loop.controlVariable}', a ${expected}, and item ${position + 1} is a number or string, no object. This will raise Run-time error '424': Object required.`, loop.sourceExpressionSpan);
			return;
		}
		if (position >= 0) {
			push('assignmentObjectTypeMismatch', `For Each Sets each item of '${loop.sourceExpression!.trim()}' into '${loop.controlVariable}', a ${expected}, and item ${position + 1} is a ${held[position]}. This will raise Run-time error '13': Type mismatch.`, loop.sourceExpressionSpan);
		}
	});
}

/** Every For Each loop in a body, nested ones included. */
/**
 * The host type For Each hands out over a host collection: its Item's type
 * (a Worksheet over Worksheets, a Workbook over Workbooks), and a Range over a
 * Range. Undefined where that is Object or Variant, as over Sheets, which
 * holds charts too.
 */
function hostElementType(
	source: string,
	span: { start: number; end: number },
	env: ReadonlyMap<string, string>,
	signatures: ReturnType<typeof buildModuleTypeSignatures>,
	sourceNames: ReturnType<typeof sourceNameScopeFor>,
	memberCtx: MemberCompletionContext,
): string | undefined {
	const toks = rawExpressionTokens(source.slice(span.start, span.end)).filter((tok) => tok.kind !== 'comment');
	// The analyzer's default host is Excel: with no model given, its globals still resolve.
	const model = memberCtx.model ?? getExcelObjectModel();
	// The model's keys keep their case: Excel.Worksheets, not excel.worksheets.
	const collection = inferExpressionType(toks, 0, env, signatures, sourceNames, source, { ...memberCtx, model })?.type;
	if (!collection || !collection.includes('.')) {
		return undefined;
	}
	if (normalizeType(collection) === 'excel.range') {
		return 'Excel.Range';
	}
	const item = getHostMembers(collection, model).find((member) => member.name === 'Item');
	const type = item?.returns;
	return type && type.includes('.') ? type : undefined;
}

function forEachLoopIn(body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined, visit: (loop: ForBlockNode) => void): void {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (node.kind === 'ForBlock' && node.each) {
			visit(node);
		}
		if ('body' in node && Array.isArray(node.body)) {
			forEachLoopIn(node.body as BodyNode[], activity, visit);
		}
	}
}

function knownClassNamed(name: string | undefined, memberCtx: MemberCompletionContext): KnownClass | undefined {
	if (!name) {
		return undefined;
	}
	if (name.toLowerCase() === 'collection') {
		return { display: 'Collection', members: COLLECTION_MEMBERS, params: COLLECTION_PARAMS };
	}
	const projectType = (memberCtx.projectClassMembers ?? []).find(
		(type) => type.kind === 'class' && type.exhaustive === true && type.name.toLowerCase() === name.toLowerCase(),
	);
	if (!projectType) {
		return undefined;
	}
	const properties = projectType.members.filter((m) => m.kind === 'property' && m.signature !== undefined);
	const params = new Map<string, readonly KnownParam[]>();
	for (const m of projectType.members) {
		// A Property Get's parameters too: `o.Idx` with Idx(ByVal i As Long) raises 449 (issue #685).
		const list = (m.kind === 'method' || (m.kind === 'property' && !m.letAccessor && !m.setAccessor)) && m.signature ? signatureParams(m.signature) : undefined;
		if (list) {
			params.set(m.name.toLowerCase(), list);
		}
	}
	return {
		params,
		display: projectType.name,
		members: new Set(projectType.members.map((m) => m.name.toLowerCase())),
		readOnly: new Set(properties.filter((m) => !m.letAccessor && !m.setAccessor).map((m) => m.name.toLowerCase())),
		writeOnly: new Set(projectType.members.filter((m) => m.kind === 'property' && m.letAccessor && m.signature === undefined).map((m) => m.name.toLowerCase())),
		setOnly: new Set(projectType.members.filter((m) => m.kind === 'property' && m.setAccessor && !m.letAccessor && m.signature === undefined).map((m) => m.name.toLowerCase())),
		noLet: new Set(projectType.members.filter((m) => m.kind === 'property' && m.setAccessor && !m.letAccessor && m.signature !== undefined).map((m) => m.name.toLowerCase())),
		subs: new Set(projectType.members.filter((m) => m.kind === 'method' && m.sub).map((m) => m.name.toLowerCase())),
		scalarFields: new Map(projectType.members
			.filter((m) => m.kind === 'property' && m.signature === undefined && !m.letAccessor && !m.setAccessor && m.returns !== undefined && SCALAR_FIELD_TYPES.has(m.returns.toLowerCase()))
			.map((m) => [m.name.toLowerCase(), m.returns!])),
	};
}

/**
 * The members of Excel's WorksheetFunction, lowercased. The list is the type
 * library's, the one Application's check already reads as complete: Ifs,
 * Switch, VStack and TextSplit are absent from both, and raise 438 through
 * either (issue #442, measured in Excel 16.0 build 20430).
 */
function worksheetFunctionNames(model: HostObjectModel | undefined): ReadonlySet<string> {
	return new Set(getHostMembers('Excel.WorksheetFunction', model).map((member) => member.name.toLowerCase()));
}

/** Excel's Application members plus the worksheet functions it also answers to. */
function excelApplicationSurface(model: HostObjectModel | undefined): ReadonlySet<string> | undefined {
	if (model && model.hostName !== undefined && model.hostName !== 'Excel') {
		return undefined;
	}
	if (getHostType('Excel.Application', model)?.exhaustive !== true) {
		return undefined;
	}
	const names = new Set<string>();
	for (const member of getHostMembers('Excel.Application', model)) {
		names.add(member.name.toLowerCase());
	}
	for (const member of getHostMembers('Excel.WorksheetFunction', model)) {
		names.add(member.name.toLowerCase());
	}
	return names;
}

/** A Range's members, when the model knows all of them; empty otherwise. */
function excelRangeSurface(model: HostObjectModel | undefined): ReadonlySet<string> {
	return getHostType('Excel.Range', model)?.exhaustive === true
		? new Set(getHostMembers('Excel.Range', model).map((member) => member.name.toLowerCase()))
		: new Set();
}

function checkStatement(
	source: string,
	base: number,
	toks: readonly VbaToken[],
	held: ReadonlyMap<string, KnownClass>,
	applicationSurface: ReadonlySet<string> | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	for (let i = 0; i + 2 < toks.length; i++) {
		// `WorksheetFunction.Mid`, `Application.WorksheetFunction.Summ`: the
		// VBE compiles any name there too, and one that is no worksheet
		// function raises 438 (issue #442, measured in Excel 16.0).
		if (applicationSurface && tokenText(toks[i]) === 'worksheetfunction' && toks[i + 1].rawText === '.') {
			const name = tokenName(toks[i + 2]);
			const functions = name ? worksheetFunctionNames(memberCtx.model) : undefined;
			if (name && functions && !functions.has(name.toLowerCase())
				&& resolveReceiverTypeAt(source, base + toks[i + 1].end, memberCtx) === 'Excel.WorksheetFunction') {
				push('runtimeMemberNotFound', `WorksheetFunction has no function '${name}'. The VBE compiles the name; this will raise Run-time error '438': Object doesn't support this property or method.`, { start: base + toks[i + 2].start, end: base + toks[i + 2].end });
			}
			continue;
		}
		if (toks[i + 1].rawText !== '.' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const receiver = tokenName(toks[i]);
		const memberName = tokenName(toks[i + 2]);
		if (!receiver || !memberName) {
			continue;
		}
		const at = { start: base + toks[i + 2].start, end: base + toks[i + 2].end };
		const known = held.get(receiver.toLowerCase());
		if (known) {
			const lower = memberName.toLowerCase();
			const nothing = known.mayBeNothing ? `, or '91' while '${receiver}' is Nothing` : '';
			if (!known.members.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, which has no member '${memberName}'. This will raise Run-time error '438': Object doesn't support this property or method${nothing}.`, at);
				continue;
			}
			// A Sub assigned, `o.M = 5`, raises 450, and one read with
			// arguments, `x = o.M(1)`, 451 (issue #414, measured in Excel 16.0).
			const statementHead = tokenText(toks[0]);
			const target = toks[i + 3]?.rawText === '=' && (i === 0 || (i === 1 && (statementHead === 'set' || statementHead === 'let')));
			if (known.subs?.has(lower)) {
				if (target) {
					push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' is a Sub, which takes no assignment. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment${nothing}.`, at);
					continue;
				}
				if (i > 0 && toks[i + 3]?.rawText === '(' && statementHead !== 'call') {
					push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' is a Sub, which gives no value to read. This will raise Run-time error '451': Property let procedure not defined and property get procedure did not return an object${nothing}.`, at);
					continue;
				}
			}
			// `o.S.Add 1` with S a String field (issue #414).
			const scalarType = known.scalarFields?.get(lower);
			if (scalarType && toks[i + 3]?.rawText === '.' && tokenName(toks[i + 4])) {
				push('variantValueMisuse', `'${receiver}' holds a ${known.display} here, whose '${memberName}' is a ${scalarType}, which has no members. This will raise Run-time error '424': Object required${nothing}.`, at);
				continue;
			}
			// `o.O = New Collection` with no Set, and O a Get and a Set (issue
			// #685, measured in Excel 16.0).
			if (target && statementHead !== 'set' && known.noLet?.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' has a Property Get and a Property Set and no Property Let, so it takes no value without Set. This will raise Run-time error '438': Object doesn't support this property or method${nothing}.`, at);
				continue;
			}
			if (!target && known.setOnly?.has(lower) && !(i > 0 && toks[i - 1]?.rawText === '.')) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' has a Property Set and no Property Get, so it has no value to read. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment${nothing}.`, at);
				continue;
			}
			// The arguments the member refuses (issue #485, measured in Excel 16.0).
			const params = known.params?.get(lower);
			const refusal = params ? argumentRefusal(toks, i, params, memberName, known.display === 'Collection' && lower === 'count') : undefined;
			if (refusal) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here: ${refusal}${nothing}.`, at);
				continue;
			}
			// `o.RO = 5` as the statement, a Let into a Get-only property.
			const assigned = i === 0 && toks[i + 3]?.rawText === '=';
			if (assigned && known.readOnly?.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' has a Property Get and no Property Let. This will raise Run-time error '451': Property let procedure not defined and property get procedure did not return an object${nothing}.`, at);
			} else if (!assigned && known.writeOnly?.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' has a Property Let and no Property Get, so it has no value to read. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment${nothing}.`, at);
			}
			continue;
		}
		if (applicationSurface && receiver.toLowerCase() === 'application' && !applicationSurface.has(memberName.toLowerCase())) {
			const type = resolveReceiverTypeAt(source, base + toks[i + 1].end, memberCtx);
			if (type === 'Excel.Application') {
				push('runtimeMemberNotFound', `Application has no member '${memberName}', and it is not a worksheet function either. The VBE compiles the name because Application is extensible; this will raise Run-time error '438': Object doesn't support this property or method.`, at);
			}
		}
	}
}

/** The scalar types a field may be declared as, whose value has no members. */
const SCALAR_FIELD_TYPES: ReadonlySet<string> = new Set(['string', 'long', 'integer', 'double', 'single', 'boolean', 'date', 'currency', 'byte', 'longlong']);

/** Collection's members, its hidden enumerator included. */
const COLLECTION_SURFACE: ReadonlySet<string> = new Set([...COLLECTION_MEMBERS, '_newenum']);

const MEMBER_NOT_SUPPORTED = 'This will raise Run-time error \'438\': Object doesn\'t support this property or method.';

/**
 * A member no early-bound receiver of an open type has (issue #305, each
 * measured in Excel 16.0): the VBE compiles the name, since the interface is
 * extensible, and the call raises 438. A local As Collection, a Range by
 * any route (`Cells.Nope`, `Range("A1").Nope`, a Range variable), a
 * variable As Application, and ActiveSheet when neither a Worksheet nor a
 * Chart nor any document module of the project has the name.
 */
function checkOpenTypeMembers(
	source: string,
	base: number,
	toks: readonly VbaToken[],
	env: ReadonlyMap<string, string>,
	applicationSurface: ReadonlySet<string> | undefined,
	rangeNames: ReadonlySet<string> | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const model = memberCtx.model;
	const projectTypes = memberCtx.projectClassMembers ?? [];
	let sheetNames: ReadonlySet<string> | undefined;
	for (let i = 1; i + 1 < toks.length; i++) {
		const name = toks[i].rawText === '.' ? tokenName(toks[i + 1]) : undefined;
		if (!name || (tokenName(toks[i - 1]) === undefined && toks[i - 1].rawText !== ')') || toks[i - 1].kind === 'keyword') {
			continue;
		}
		const lower = name.toLowerCase();
		const at = { start: base + toks[i + 1].start, end: base + toks[i + 1].end };
		// A plain name before the dot: `c.Nope`, `a.Nope`, `ActiveSheet.Nope`.
		const receiver = toks[i - 1].rawText !== ')' && toks[i - 2]?.rawText !== '.' ? tokenName(toks[i - 1]) : undefined;
		const declared = receiver ? normalizeType(env.get(receiver.toLowerCase())) : undefined;
		if (declared === 'collection' || declared === 'vba.collection') {
			if (!COLLECTION_SURFACE.has(lower) && !projectTypes.some((type) => type.name.toLowerCase() === 'collection')) {
				push('runtimeMemberNotFound', `'${receiver}' is a Collection, which has only Add, Count, Item and Remove. The VBE compiles '${name}'; ${MEMBER_NOT_SUPPORTED}`, at);
			}
			continue;
		}
		if (!applicationSurface) {
			continue;
		}
		if ((declared === 'application' || declared === 'excel.application') && !applicationSurface.has(lower)) {
			push('runtimeMemberNotFound', `'${receiver}' is an Application, which has no member '${name}', and it is not a worksheet function either. The VBE compiles the name because Application is extensible; ${MEMBER_NOT_SUPPORTED}`, at);
			continue;
		}
		if (receiver && tokenText(toks[i - 1]) === 'activesheet' && !env.has('activesheet')) {
			sheetNames ??= sheetSurface(model, projectTypes);
			if (!sheetNames.has(lower)) {
				push('runtimeMemberNotFound', `ActiveSheet has no member '${name}': neither a Worksheet nor a Chart has one, and no document module of the project declares it. ${MEMBER_NOT_SUPPORTED}`, at);
			}
			continue;
		}
		if (rangeNames && rangeNames.size > 0 && !rangeNames.has(lower) && resolveReceiverTypeAt(source, base + toks[i].end, memberCtx) === 'Excel.Range') {
			push('runtimeMemberNotFound', `A Range has no member '${name}'. The VBE compiles the name because Range is extensible; ${MEMBER_NOT_SUPPORTED}`, at);
		}
	}
}

/** What ActiveSheet may answer to: a Worksheet's members, a Chart's, and every document module's. */
function sheetSurface(model: HostObjectModel | undefined, projectTypes: NonNullable<MemberCompletionContext['projectClassMembers']>): ReadonlySet<string> {
	const names = new Set<string>();
	for (const type of ['Excel.Worksheet', 'Excel.Chart']) {
		for (const member of getHostMembers(type, model)) {
			names.add(member.name.toLowerCase());
		}
	}
	for (const type of projectTypes) {
		if (type.kind === 'document') {
			for (const member of type.members) {
				names.add(member.name.toLowerCase());
			}
		}
	}
	return names;
}

/**
 * `f.Controls("Nope")` on a form whose controls are known, with no control of
 * that name (case-insensitive, those inside a Frame included), raises
 * -2147024809, "Could not find the specified object" (issue #226, measured in
 * Excel 16.0). `Me.Controls(...)` inside the form does the same.
 */
function checkFormControlNames(
	source: string,
	base: number,
	toks: readonly VbaToken[],
	memberCtx: MemberCompletionContext,
	added: ReadonlySet<string> | 'any',
	push: PushFn,
): void {
	for (let i = 1; i + 3 < toks.length; i++) {
		if (tokenText(toks[i]) !== 'controls' || toks[i - 1].rawText !== '.' || toks[i + 1].rawText !== '('
			|| (toks[i + 2].kind !== 'stringLiteral' && toks[i + 2].kind !== 'integerLiteral') || toks[i + 3].rawText !== ')') {
			continue;
		}
		const form = projectTypeAt(source, base + toks[i - 1].end, memberCtx);
		if (form?.kind !== 'userform' || form.exhaustive !== true || added === 'any') {
			continue;
		}
		const controls = form.members.filter((member) => /^MSForms\./i.test(member.returns ?? ''));
		// `Me.Controls(99)`: Controls counts from 0 (issue #315, measured in
		// Excel 16.0). A procedure that adds a control is not judged.
		if (toks[i + 2].kind === 'integerLiteral') {
			const index = Number(toks[i + 2].rawText);
			if (added.size === 0 && index >= controls.length) {
				push(
					'runtimeMemberNotFound',
					`The form ${form.name} has ${controls.length} control${controls.length === 1 ? '' : 's'}, indexed 0 to ${controls.length - 1}; ${index} is none of them. This will raise Run-time error '-2147024809': Invalid argument.`,
					{ start: base + toks[i + 2].start, end: base + toks[i + 2].end },
				);
			}
			continue;
		}
		const name = stringLiteralValue(toks[i + 2].rawText);
		// `Me.Controls.Add "Forms.TextBox.1", "Dyn"` names one the designer lacks.
		if (added.has(name.toLowerCase())) {
			continue;
		}
		if (!controls.some((control) => control.name.toLowerCase() === name.toLowerCase())) {
			push(
				'runtimeMemberNotFound',
				`The form ${form.name} has no control named "${name}". This will raise Run-time error '-2147024809': Could not find the specified object.`,
				{ start: base + toks[i + 2].start, end: base + toks[i + 2].end },
			);
		}
	}
}

/**
 * The control names a procedure gives `Controls.Add` as a literal second
 * argument, lowercased (issue #315), or 'any' when one is added under a
 * name the code does not spell out.
 */
function controlsAddedIn(source: string, body: BodyNode[], activity: ConditionalActivityTracker | undefined): ReadonlySet<string> | 'any' {
	const names = new Set<string>();
	let any = false;
	forEachStatement(body, (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			const toks = statementTokens(source, span).filter((tok) => tok.kind !== 'comment');
			for (let i = 2; i < toks.length; i++) {
				if (tokenText(toks[i]) !== 'add' || toks[i - 1].rawText !== '.' || tokenText(toks[i - 2]) !== 'controls') {
					continue;
				}
				const open = toks[i + 1]?.rawText === '(' ? i + 1 : -1;
				const close = open > 0 ? matchParenFrom([...toks], open) : toks.length;
				const args = splitTopLevelTokenGroups([...toks], open > 0 ? open + 1 : i + 1, ',', close);
				const named = args.find((arg) => arg[1]?.rawText === ':=' && tokenText(arg[0]) === 'name');
				const arg = named ? named.slice(2) : args[1]?.[1]?.rawText === ':=' ? undefined : args[1];
				if (arg?.length === 1 && arg[0].kind === 'stringLiteral') {
					names.add(stringLiteralValue(arg[0].rawText).toLowerCase());
				} else {
					any = true;
				}
			}
		}
	}, activity);
	return any ? 'any' : names;
}

/** A tracked variable named in any position other than `name.Member` is no longer followed. */
function forgetOtherUses(toks: readonly VbaToken[], held: Map<string, KnownClass>): void {
	for (let i = 0; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		if (lower && held.has(lower) && toks[i - 1]?.rawText !== '.' && toks[i + 1]?.rawText !== '.') {
			held.delete(lower);
		}
	}
}

function forgetMentioned(toks: readonly VbaToken[], held: Map<string, KnownClass>): void {
	for (const tok of toks) {
		const lower = tokenName(tok)?.toLowerCase();
		if (lower && held.has(lower)) {
			held.delete(lower);
		}
	}
}
