// Rule family: declaration-site rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: procedure headers, identifier
// spelling, module-level declaration placement, reserved names, property
// accessor signatures, parameter order/defaults, Dim initializers, As-clause
// and fixed-length-string declarations, and Option placement.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import {
	isCreatableTypeCompletion,
	resolveTypeName,
	type TypeCompletionKind,
} from '../../completion/typeCompletion';
import {
	collectConditionalDirectives,
	type ConditionalActivityTracker,
} from '../../conditional/conditionalCompilation';
import { isReservedIdentifier, OPERATOR_IDENTIFIERS } from '../../lexer/keywordTable';
import type { VbaToken } from '../../lexer/tokenKinds';
import { parseFixedLengthStringType } from '../../parser/fixedLengthString';
import type {
	BodyNode,
	ModuleMember,
	ModuleNode,
	ParameterNode,
	ProcedureNode,
	Span,
	LeafStatementNode,
	TypeFieldNode,
	VariableDeclNode,
	VariableGroupNode,
} from '../../parser/nodes';
import { isTypeDeclarationSuffix } from '../../parser/typeDeclarationSuffix';
import { resolveRuntimeFunction } from '../../runtime/vbaRuntime';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import {
	collectTypeNameReferences,
	type TypeNameReferenceKind,
	typeReferenceLookupName,
} from '../../semantic/typeSemanticTokens';
import {
	type AnalyzeModuleOptions,
	isObjectModuleKind,
	type PushFn,
} from '../analysisContext';
import type { InferredArgumentType } from '../callExtraction';
import { juxtaposedValueIndex } from './expressions';
import {
	collectBodyLiteralIntegerConstants,
	collectModuleLiteralIntegerConstants,
	resolveFixedLengthStringSize,
} from '../constExpr';
import { libraryTypeNames } from '../../host/libraryTypeNames';
import {
	declarationNameHit,
	DEFTYPE_KEYWORDS,
	leadingDeclarationModifierCount,
	moduleDeclarationStatementInProcedure,
	NameTokenHit,
	nameTokenHit,
	reportRepeatedKeys,
	scanConditionalCompilationBranchOrder,
} from '../rules/shared';
import {
	incompatibilityReason,
	inferArgumentType,
	isKnownScalarType,
	normalizeType,
	createObjectAssignmentTypeResolver,
	spanForTokens,
} from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	declaredNameSpan,
	firstTokenSpan,
	forEachStatement,
	forEachVariableGroup,
	isInactiveNode,
	matchParenFrom,
	pluralizeCount,
	statementTokens,
	statementTokensAfterLeadingLabel,
	stripHeaderBrackets,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
} from '../walker';

/** Access/storage modifiers that may lead a procedure declaration. */
const PROC_MODIFIERS = new Set([
	'public', 'private', 'friend', 'global', 'static',
]);

/**
 * Rule: a procedure header must be `[(modifiers)] Sub|Function|Property Get/Let/Set
 * Name [(params)] [As Type]`. Once the name is read, the only legal next token is
 * `(` (the parameter list) or, for a `Function`/`Property Get`, `As` (the return
 * type). Any other token - most commonly a second word, as in `Sub My Sub`, where
 * the name was meant to contain a space - is the VBE "Expected: (" compile error.
 * Property `Let`/`Set` and `Sub` have no return value, so an `As` right after the
 * name is rejected for them too.
 */
export function checkProcedureHeader(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const headerStart = member.span.start;
		const nl = source.indexOf('\n', headerStart);
		const headerEnd = nl === -1 ? member.span.end : nl;
		const toks = statementTokens(source, { start: headerStart, end: headerEnd });

		let i = 0;
		while (i < toks.length && PROC_MODIFIERS.has(toks[i].rawText.toLowerCase())) {
			i++;
		}
		const kw = toks[i]?.rawText.toLowerCase();
		let allowAs = false;
		if (kw === 'function') {
			allowAs = true;
			i++;
		} else if (kw === 'sub') {
			i++;
		} else if (kw === 'property') {
			i++;
			if (toks[i]?.rawText.toLowerCase() === 'get') {
				allowAs = true;
			}
			i++; // skip the accessor (Get/Let/Set)
		} else {
			continue; // not a recognised procedure header
		}

		const nameTok = toks[i];
		if (!nameTok) {
			continue; // malformed in a way the structural analyzer already reports
		}
		if (isDigitStartedToken(nameTok)) {
			continue; // invalid-identifier-start owns the precise declaration-name range
		}
		let nextIndex = i + 1;
		if (
			allowAs &&
			toks[nextIndex] &&
			nameTok.end === toks[nextIndex].start &&
			isTypeDeclarationSuffix(toks[nextIndex].rawText)
		) {
			nextIndex++;
		}
		const next = toks[nextIndex];
		if (!next) {
			continue; // `Sub Foo` with no parameter list is legal
		}
		const r = next.rawText;
		if (r === '(' || (allowAs && r.toLowerCase() === 'as')) {
			continue;
		}
		push(
			'invalidProcedureHeader',
			`Unexpected '${r}' after procedure name '${stripHeaderBrackets(nameTok.rawText)}'; a procedure name must be a single identifier.`,
			{ start: headerStart + next.start, end: headerStart + next.end },
		);
	}
}

/** Strips the surrounding `[ ]` from a bracketed identifier, if present. */
export function checkInvalidIdentifierStarts(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (kind: string, hit: InvalidIdentifierStartHit | undefined): void => {
		if (!hit) {
			return;
		}
		if (hit.reason === 'digit') {
			push(
				'invalidIdentifierStart',
				`Invalid ${kind} name '${hit.name}': identifiers cannot start with a digit.`,
				hit.span,
			);
		} else if (hit.reason === 'underscore') {
			push(
				'invalidIdentifierStart',
				`Invalid ${kind} name '${hit.name}': identifiers cannot start with an underscore.`,
				hit.span,
			);
		} else {
			push(
				'invalidIdentifierCharacter',
				`Invalid ${kind} name '${hit.name}': '${hit.reason === 'hyphen' ? '-' : '.'}' is not allowed in an identifier.`,
				hit.span,
			);
		}
	};

	const inspectVariableGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			report('variable', invalidDeclarationIdentifierStart(source, decl.span));
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectVariableGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			report('user-defined type', invalidTypeOrEnumIdentifierStart(source, member.span, 'type'));
			for (const field of member.fields) {
				report('type field', invalidDeclarationIdentifierStart(source, field.span));
			}
			continue;
		}
		if (member.kind === 'Enum') {
			report('enum', invalidTypeOrEnumIdentifierStart(source, member.span, 'enum'));
			for (const enumMember of member.members) {
				report('enum member', invalidDeclarationIdentifierStart(source, enumMember.span));
			}
			continue;
		}
		if (member.kind === 'Declare') {
			report('Declare procedure', invalidDeclareIdentifierStart(source, member.span));
			continue;
		}
		if (member.kind === 'ConditionalDirective') {
			report('conditional compiler constant', invalidConstDirectiveIdentifierStart(source, member.span));
			continue;
		}
		if (member.kind !== 'Procedure') {
			continue;
		}
		report('procedure', invalidProcedureIdentifierStart(source, member));
		for (const param of member.params) {
			report('parameter', invalidParameterIdentifierStart(source, param.span));
		}
		forEachVariableGroup(member.body, inspectVariableGroup, activity);
	}
}

interface InvalidIdentifierStartHit {
	name: string;
	span: Span;
	reason: 'digit' | 'underscore' | 'hyphen' | 'dot';
}

function invalidDeclarationIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	return invalidIdentifierStartAt(source, span, toks, 0);
}

function invalidParameterIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	let i = 0;
	while (isParameterModifier(toks[i])) {
		i++;
	}
	return invalidIdentifierStartAt(source, span, toks, i);
}

function invalidProcedureIdentifierStart(
	source: string,
	proc: ProcedureNode,
): InvalidIdentifierStartHit | undefined {
	const header = firstLineSpan(source, proc.span);
	const toks = statementTokens(source, header);
	let i = 0;
	while (i < toks.length && PROC_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	const head = tokenText(toks[i]);
	if (head === 'property') {
		i += 2;
	} else if (head === 'sub' || head === 'function') {
		i++;
	}
	return invalidIdentifierStartAt(source, header, toks, i);
}

function invalidTypeOrEnumIdentifierStart(
	source: string,
	span: Span,
	keyword: 'type' | 'enum',
): InvalidIdentifierStartHit | undefined {
	const header = firstLineSpan(source, span);
	const toks = statementTokens(source, header);
	let i = 0;
	if (tokenText(toks[i]) === 'public' || tokenText(toks[i]) === 'private') {
		i++;
	}
	if (tokenText(toks[i]) === keyword) {
		i++;
	}
	return invalidIdentifierStartAt(source, header, toks, i);
}

function invalidDeclareIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	const kindIndex = toks.findIndex(
		(tok) => tokenText(tok) === 'sub' || tokenText(tok) === 'function',
	);
	return invalidIdentifierStartAt(source, span, toks, kindIndex + 1);
}

function invalidConstDirectiveIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	return tokenText(toks[1]) === 'const'
		? invalidIdentifierStartAt(source, span, toks, 2)
		: undefined;
}

function invalidIdentifierStartAt(
	source: string,
	base: Span,
	toks: readonly VbaToken[],
	index: number,
): InvalidIdentifierStartHit | undefined {
	const tok = toks[index];
	if (!tok || tok.kind === 'bracketedIdentifier') {
		return undefined; // [bracketed] names may contain anything
	}
	// Embedded invalid character: a name token directly followed by '-' or '.'
	// (e.g. `user-name`, `bad.name`). The parser keeps the first token as the name
	// and leaves the rest, so the malformation is only visible in the token stream.
	const next = toks[index + 1];
	if (tok.kind === 'identifier' && (next?.rawText === '-' || next?.rawText === '.')) {
		const start = base.start + tok.start;
		const after = toks[index + 2];
		const end = base.start + (after ? after.end : next.end);
		return {
			name: source.slice(start, end),
			span: { start, end },
			reason: next.rawText === '-' ? 'hyphen' : 'dot',
		};
	}
	// Invalid start character: a digit or a leading underscore.
	let reason: InvalidIdentifierStartHit['reason'] | undefined;
	if (isDigitStartedToken(tok)) {
		reason = 'digit';
	} else if (tok.rawText.startsWith('_')) {
		reason = 'underscore';
	}
	if (!reason) {
		return undefined;
	}
	const start = base.start + tok.start;
	const end = invalidIdentifierTextEnd(source, start, base.end);
	return { name: source.slice(start, end), span: { start, end }, reason };
}

function isDigitStartedToken(tok: VbaToken): boolean {
	return (tok.kind === 'integerLiteral' || tok.kind === 'floatLiteral') && /^\d/.test(tok.rawText);
}

function invalidIdentifierTextEnd(source: string, start: number, limit: number): number {
	let end = start;
	while (end < limit && isInvalidIdentifierTextChar(source[end])) {
		end++;
	}
	return end;
}

function isInvalidIdentifierTextChar(ch: string | undefined): boolean {
	return ch !== undefined && /[A-Za-z0-9_]/.test(ch);
}

function isParameterModifier(tok: VbaToken | undefined): boolean {
	switch (tokenText(tok)) {
		case 'optional':
		case 'byval':
		case 'byref':
		case 'paramarray':
			return true;
		default:
			return false;
	}
}

export function checkModuleDeclarationsInProcedureBodies(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspectStatement = (stmt: LeafStatementNode): void => {
		const hit = moduleDeclarationStatementInProcedure(source, stmt.span);
		if (!hit) {
			return;
		}
		push(
			'moduleDeclarationInProcedure',
			`${hit.label} must appear in the module declarations section, not inside a procedure.`,
			hit.span,
		);
	};
	const inspectProcedureBody = (procedure: ProcedureNode): void => {
		let sawConditionalDirective = false;
		for (const node of procedure.body) {
			if (node.kind === 'ConditionalDirective') {
				sawConditionalDirective = true;
				continue;
			}
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (node.kind === 'Statement') {
				if (
					sawConditionalDirective &&
					isAlternativeProcedureHeaderStatement(source, node.span, procedure)
				) {
					continue;
				}
				inspectStatement(node);
				continue;
			}
			if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
				forEachStatement((node as { body: BodyNode[] }).body, inspectStatement, activity);
			}
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			inspectProcedureBody(member);
		}
	}
}

/**
 * Rule: module declarations belong in the declaration section before the first
 * procedure. Multiple procedures may follow each other, but once an active
 * procedure appears, later active module declarations are misplaced.
 */
export function checkModuleDeclarationsAfterProcedures(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Procedures that precede the declaration under test AND could be compiled
	// beside it. A procedure in one arm of a `#If` chain and a declaration in
	// another arm never reach the compiler together, so the declaration is not
	// "after" it in any build (issues/58).
	const proceduresAbove: Span[] = [];
	const malformedConditionalBlocks = scanConditionalCompilationBranchOrder(mod).malformedBlockSpans;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			proceduresAbove.push(member.span);
			continue;
		}
		const compiledTogether = proceduresAbove.some(
			(prior) => !activity?.mutuallyExclusive(prior, member.span),
		);
		if (!compiledTogether) {
			continue;
		}
		const hit = moduleDeclarationAfterProcedureHit(source, member);
		if (!hit) {
			continue;
		}
		if (malformedConditionalBlocks.some((span) => containsSpan(span, member.span))) {
			continue;
		}
		push(
			'moduleDeclarationAfterProcedure',
			moduleDeclarationAfterProcedureMessage(hit.label, mod, member, activity),
			hit.span,
		);
	}
}

/**
 * Rule: executable statements belong inside procedures. The module body accepts
 * declarations plus a small set of statement-shaped declaration forms (`Def*`
 * and object-module `Implements`, which has its own placement rule).
 */
export function checkModuleLevelStatementsOutsideProcedures(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Statement') {
			continue;
		}
		const hit = moduleLevelStatementOutsideProcedureHit(source, member.span);
		if (!hit) {
			continue;
		}
		push(
			'statementOutsideProcedure',
			`${hit.label} is invalid outside a Sub, Function, or Property procedure.`,
			hit.span,
		);
	}
}

function moduleLevelStatementOutsideProcedureHit(
	source: string,
	span: Span,
): { label: string; span: Span } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const head = tokenText(first);
	if (DEFTYPE_KEYWORDS.has(head) || head === 'implements') {
		return undefined;
	}
	return {
		label: `${first.canonicalText ?? first.rawText} statement`,
		span: absoluteSpan(span, first),
	};
}

function moduleDeclarationAfterProcedureMessage(
	label: string,
	mod: ModuleNode,
	member: ModuleMember,
	activity: ConditionalActivityTracker | undefined,
): string {
	if (!isInsideModuleConditionalCompilationBlock(mod, member.span)) {
		return `${label} belong in the module declarations section, before procedures.`;
	}
	const branchStatus = activity?.activityForSpan(member.span);
	if (branchStatus === 'active') {
		return `${label} in the active conditional-compilation branch belong in the module declarations section, before procedures.`;
	}
	return `${label} in a conditional-compilation branch belong in the module declarations section, before procedures.`;
}

function isInsideModuleConditionalCompilationBlock(
	mod: ModuleNode,
	span: Span,
): boolean {
	let depth = 0;
	for (const { directive, container } of collectConditionalDirectives(mod)) {
		if (container.kind !== 'module') {
			continue;
		}
		if (directive.span.start >= span.start) {
			break;
		}
		switch (directive.directiveKind) {
			case 'If':
				depth++;
				break;
			case 'EndIf':
				depth = Math.max(0, depth - 1);
				break;
			case 'Const':
			case 'ElseIf':
			case 'Else':
			case 'Unknown':
				break;
		}
	}
	return depth > 0;
}

function moduleDeclarationAfterProcedureHit(
	source: string,
	member: ModuleMember,
): { label: string; span: Span } | undefined {
	switch (member.kind) {
		case 'Declare':
			return {
				label: 'Declare statements',
				span: keywordSpan(source, member.span, 'declare'),
			};
		case 'Event':
			return {
				label: 'Event declarations',
				span: keywordSpan(source, member.span, 'event'),
			};
		case 'VariableGroup':
			return {
				label: member.isConst ? 'Const declarations' : 'Module variable declarations',
				span: member.isConst
					? keywordSpan(source, member.span, 'const')
					: firstTokenSpan(source, member.span),
			};
		case 'Type':
			return {
				label: 'Type declarations',
				span: keywordSpan(source, member.span, 'type'),
			};
		case 'Enum':
			return {
				label: 'Enum declarations',
				span: keywordSpan(source, member.span, 'enum'),
			};
		case 'Statement':
			return deftypeModuleDeclarationHit(source, member.span);
		default:
			return undefined;
	}
}

function deftypeModuleDeclarationHit(
	source: string,
	span: Span,
): { label: string; span: Span } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const first = toks[0];
	if (!first || !DEFTYPE_KEYWORDS.has(tokenText(first))) {
		return undefined;
	}
	return {
		label: `${first.canonicalText ?? first.rawText} statements`,
		span: absoluteSpan(span, first),
	};
}

function isAlternativeProcedureHeaderStatement(
	source: string,
	span: Span,
	procedure: ProcedureNode,
): boolean {
	const toks = statementTokensAfterLeadingLabel(source, span);
	let i = leadingDeclarationModifierCount(toks);
	const head = tokenText(toks[i]);
	let kind: ProcedureNode['procKind'] | undefined;
	if (head === 'property') {
		const accessor = tokenText(toks[i + 1]);
		kind =
			accessor === 'get'
				? 'PropertyGet'
				: accessor === 'let'
					? 'PropertyLet'
					: accessor === 'set'
						? 'PropertySet'
						: undefined;
		i += 2;
	} else if (head === 'function') {
		kind = 'Function';
		i += 1;
	} else if (head === 'sub') {
		kind = 'Sub';
		i += 1;
	}
	const name = tokenName(toks[i]);
	return !!kind &&
		kind === procedure.procKind &&
		!!name &&
		name.toLowerCase() === procedure.name.toLowerCase();
}

/**
 * The names the VBE refuses for a module (issues #247 and #357, measured in
 * Excel 16.0): adding one fails with 0x800AC3D4, renaming to one with
 * 50132. Line, Width, Name, Err, Mid, Time, Error, Reset, Beep, Load,
 * Unload, Access, Base, Compare, Explicit, Object, Property and Step are
 * accepted.
 */
const REFUSED_MODULE_NAMES: ReadonlySet<string> = new Set([
	'addressof', 'and', 'any', 'array', 'as', 'attribute', 'boolean', 'byref', 'byte', 'byval',
	'call', 'case', 'cdate', 'circle', 'close', 'const', 'currency', 'date', 'debug', 'decimal', 'declare', 'dim',
	'do', 'double', 'each', 'else', 'elseif', 'empty', 'end', 'enum', 'eqv', 'erase',
	'event', 'exit', 'false', 'for', 'friend', 'function', 'get', 'global', 'gosub', 'goto', 'if',
	'imp', 'implements', 'in', 'input', 'integer', 'is', 'lbound', 'len', 'lenb', 'let',
	'like', 'lock', 'long', 'longlong', 'longptr', 'loop', 'lset', 'me', 'mod', 'new', 'next', 'not',
	'nothing', 'null', 'on', 'open', 'option', 'optional', 'or', 'paramarray', 'preserve', 'print',
	'private', 'pset', 'public', 'put', 'raiseevent', 'redim', 'rem', 'resume', 'return', 'rset',
	'scale', 'seek', 'select', 'set', 'shared', 'single', 'spc', 'static', 'stop', 'string',
	'sub', 'tab', 'then', 'to', 'true', 'type', 'typeof', 'unlock', 'until', 'variant',
	'wend', 'while', 'with', 'withevents', 'write', 'xor',
]);

/**
 * The libraries every project of a host references, which a module cannot
 * share a name with: renaming one to Excel, VBA, Office or stdole fails
 * with 32813, "Name conflicts with existing module, project, or object
 * library" (issue #357, measured in Excel 16.0). Word, Access and
 * PowerPoint are accepted there, being libraries an Excel project does
 * not reference.
 */
const REFERENCED_LIBRARIES: ReadonlySet<string> = new Set(['vba', 'office', 'stdole']);
const HOST_LIBRARIES: ReadonlySet<string> = new Set(['excel', 'word', 'powerpoint', 'access']);

/**
 * A module named a word the VBE refuses. A file can still hold one, and its
 * procedures run called bare, but a call through its name does not compile.
 * The `Attribute VB_Name` line is marked, or else the first line.
 */
export function checkModuleName(source: string, moduleName: string | undefined, push: PushFn, hostName?: string): void {
	const lower = moduleName?.toLowerCase() ?? '';
	const library = REFERENCED_LIBRARIES.has(lower) || (HOST_LIBRARIES.has(lower) && lower === (hostName ?? 'Excel').toLowerCase());
	if (!moduleName || (!library && !REFUSED_MODULE_NAMES.has(lower))) {
		return;
	}
	const attribute = /^[ \t]*Attribute[ \t]+VB_Name[ \t]*=[ \t]*"([^"]*)"/im.exec(source);
	const start = attribute ? attribute.index + attribute[0].length - attribute[1].length - 2 : 0;
	const end = attribute ? start + attribute[1].length + 2 : Math.max(0, source.search(/\r?\n|$/));
	push(
		'invalidDeclarationName',
		library
			? `'${moduleName}' names an object library every project here references, so it cannot name a module: the VBE refuses the name ("Name conflicts with existing module, project, or object library").`
			: `Reserved VBA keyword '${moduleName}' cannot name a module: the VBE refuses to add one, and a call through the name, ${moduleName}.Proc, does not compile.`,
		{ start, end },
	);
}

/** True when the text after `As` names a type: `Long`, `[Long]`, `New Collection`. */
function namesAsType(asType: string | undefined): boolean {
	return /^(?:New\s+|(?!New(?:\s|$)))(?:[\p{L}_]|\[[^\]])/iu.test(asType ?? '');
}

export function checkReservedDeclarationNames(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (kind: string, hit: NameTokenHit | undefined): void => {
		if (!hit || hit.bracketed || !isReservedIdentifier(hit.name)) {
			return;
		}
		if (kind === 'type field' && hit.name.toLowerCase() === 'type') {
			return;
		}
		push(
			'invalidDeclarationName',
			`Reserved VBA keyword '${hit.name}' cannot be used as a ${kind} name.`,
			hit.span,
		);
	};

	const inspectVariableGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			report('variable', declarationNameHit(source, decl.span, decl.name));
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectVariableGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			report('user-defined type', typeOrEnumNameHit(source, member.span, 'type'));
			for (const field of member.fields) {
				// A reserved word names a Type member when an As clause with a type
				// follows: `Next As Long`, `Long As Long`, `Loop As New Collection`
				// (VBE oracle reserved_member_name_*). An array clause, a type
				// suffix or an As with no type is still a syntax error, and so is
				// `Me As Long` (Expected: =).
				if (field.name.toLowerCase() !== 'me' && field.hasAsClause && !field.typeSuffix && !field.isArray && namesAsType(field.asType)) {
					continue;
				}
				report('type field', declarationNameHit(source, field.span, field.name));
			}
			continue;
		}
		if (member.kind === 'Enum') {
			report('enum', typeOrEnumNameHit(source, member.span, 'enum'));
			for (const enumMember of member.members) {
				report('enum member', declarationNameHit(source, enumMember.span, enumMember.name));
			}
			continue;
		}
		if (member.kind === 'Declare') {
			report('Declare procedure', declareNameHit(source, member.span));
			continue;
		}
		if (member.kind !== 'Procedure') {
			continue;
		}
		report('procedure', procedureNameHit(source, member));
		for (const param of member.params) {
			report('parameter', declarationNameHit(source, param.span, param.name));
		}
		forEachVariableGroup(member.body, inspectVariableGroup, activity);
	}
}

/**
 * Rule: Property Let/Set setters receive the assigned value through the final
 * parameter. A setter with no parameters has no value slot, setters have no
 * return type, and Property Set value parameters must be object references.
 * A Property Let's value parameter may be of any type (issue #107).
 */
export function checkPropertySetterValueParameters(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'PropertyLet' && member.procKind !== 'PropertySet')
		) {
			continue;
		}
		if (member.hasAsClause) {
			const label = member.procKind === 'PropertyLet' ? 'Property Let' : 'Property Set';
			push(
				'propertySetterReturnType',
				`${label} '${member.name}' cannot declare a return type; use the final value parameter for the assigned value.`,
				propertySetterReturnTypeSpan(source, member),
			);
		}
		// A ParamArray cannot be the value: `Property Let P(ParamArray v())` is
		// "Argument not optional" (issue #266, measured in Excel 16.0).
		if (member.params.length > 0 && !member.params[member.params.length - 1].paramArray) {
			const valueParam = member.params[member.params.length - 1];
			if (member.procKind === 'PropertySet') {
				const normalized = normalizeType(valueParam.asType);
				if (normalized && isKnownScalarType(normalized)) {
					push(
						'propertySetScalarValue',
						`Property Set '${member.name}' final value parameter '${valueParam.name}' must be an object reference, but it is declared As ${valueParam.asType}.`,
						declaredNameSpan(source, valueParam.span, valueParam.name),
					);
				}
			}
			// A Property Let's value parameter may be any type, object types
			// included: `Property Let Item(ByVal v As Object)`, `As Worksheet` and
			// `As <project class>` all compile, and `h.Item = New Collection`
			// calls the Let (issue #107, measured in Excel 16.0). The old
			// property-let-object-value report was wrong and is retired.
			continue;
		}
		const label = member.procKind === 'PropertyLet' ? 'Property Let' : 'Property Set';
		push(
			'propertySetterMissingValue',
			member.params.length > 0
				? `${label} '${member.name}' must take its value in a parameter after its ParamArray. This is a VBE compile error: Argument not optional.`
				: `${label} '${member.name}' must include a final value parameter.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function propertySetterReturnTypeSpan(source: string, proc: ProcedureNode): Span {
	const header = firstLineSpan(source, proc.span);
	const toks = statementTokens(source, header);
	let i = 0;
	while (i < toks.length && PROC_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	if (tokenText(toks[i]) === 'property') {
		i += 2; // Property + Let/Set
	}
	i++; // property name
	if (toks[i]?.rawText !== '(') {
		return keywordSpan(source, header, 'as');
	}
	let depth = 0;
	while (i < toks.length) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
			if (depth === 0) {
				i++;
				break;
			}
		}
		i++;
	}
	if (tokenText(toks[i]) !== 'as') {
		return keywordSpan(source, header, 'as');
	}
	const asToken = toks[i];
	const typeStart = i + 1;
	let typeEnd = consumeDeclarationTypeName(toks, typeStart);
	if (typeEnd === typeStart) {
		typeEnd = i + 1;
	}
	const endToken = toks[typeEnd - 1] ?? asToken;
	return {
		start: header.start + asToken.start,
		end: header.start + endToken.end,
	};
}

interface PropertyAccessorGroup {
	name: string;
	gets: ProcedureNode[];
	setters: ProcedureNode[];
}

/**
 * Rule: paired Property Get and Let/Set declarations for the same property use
 * the same index-argument shape. Let/Set add a final assigned-value parameter,
 * which is not part of the index-argument comparison.
 */
export function checkPropertyAccessorSignatures(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const groups = new Map<string, PropertyAccessorGroup>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'PropertyGet' &&
				member.procKind !== 'PropertyLet' &&
				member.procKind !== 'PropertySet')
		) {
			continue;
		}
		const key = member.name.toLowerCase();
		let group = groups.get(key);
		if (!group) {
			group = { name: member.name, gets: [], setters: [] };
			groups.set(key, group);
		}
		if (member.procKind === 'PropertyGet') {
			group.gets.push(member);
		} else {
			group.setters.push(member);
		}
	}

	for (const group of groups.values()) {
		if (group.gets.length === 0 && group.setters.length === 2) {
			// A Let and a Set with no Get: their indexes keep one set of names
			// (issue #266, measured in Excel 16.0).
			const [first, second] = group.setters;
			const firstIndexes = first.params.slice(0, -1);
			const secondIndexes = second.params.slice(0, -1);
			for (let i = 0; i < Math.min(firstIndexes.length, secondIndexes.length) && firstIndexes.length === secondIndexes.length; i++) {
				const reason = propertyParameterNameMismatch(firstIndexes[i], secondIndexes[i], i + 1);
				if (reason) {
					push(
						'propertyAccessorSignatureMismatch',
						`${propertyProcedureLabel(second.procKind)} '${second.name}' argument list must match ${propertyProcedureLabel(first.procKind)} '${first.name}' before the final value parameter. ${reason}`,
						declaredNameSpan(source, second.span, second.name),
					);
					break;
				}
			}
			continue;
		}
		if (group.gets.length !== 1) {
			continue;
		}
		const getter = group.gets[0];
		for (const setter of group.setters) {
			if (setter.params.length === 0) {
				continue;
			}
			const reason = propertyIndexParameterMismatch(
				getter.params,
				setter.params.slice(0, -1),
			);
			if (reason) {
				push(
					'propertyAccessorSignatureMismatch',
					`${propertyProcedureLabel(setter.procKind)} '${setter.name}' argument list must match Property Get '${getter.name}' before the final value parameter. ${reason}`,
					declaredNameSpan(source, setter.span, setter.name),
				);
				continue;
			}
			// A Let's value parameter must have the Get's type: `Get Size() As
			// Long` with `Let Size(ByVal v As Integer)` is "Definitions of
			// property procedures for the same property are inconsistent"
			// (issue #124, measured in Excel 16.0). Either side without a type
			// is Variant. A Set's value is never compared: Variant, Object,
			// Collection and even Long Gets beside an Object or Collection Set
			// all compile (issue #152, measured).
			if (setter.procKind !== 'PropertyLet') {
				continue;
			}
			const valueParam = setter.params[setter.params.length - 1];
			const getType = normalizeType(getter.returnType) ?? (getter.typeSuffix ? undefined : 'variant');
			const valueType = normalizeType(valueParam.asType) ?? (valueParam.typeSuffix ? undefined : 'variant');
			// An object Get beside a Variant Let compiles: `Get M() As Collection`
			// with `Let M(ByVal v As Variant)` (issue #414, measured in Excel
			// 16.0). A Collection Get with an Object Let does not.
			const objectGetVariantLet = valueType === 'variant' && getType !== undefined && getType !== 'variant' && !isKnownScalarType(getType);
			if (getType !== undefined && valueType !== undefined && getType !== valueType && !valueParam.isArray && !objectGetVariantLet) {
				push(
					'propertyAccessorSignatureMismatch',
					`${propertyProcedureLabel(setter.procKind)} '${setter.name}' takes its value As ${valueParam.asType ?? 'Variant'}, but Property Get '${getter.name}' returns ${getter.returnType ?? 'Variant'}; the definitions of a property's procedures must agree.`,
					declaredNameSpan(source, valueParam.span, valueParam.name),
				);
			}
		}
	}
}

function propertyIndexParameterMismatch(
	getParams: readonly ParameterNode[],
	setterIndexParams: readonly ParameterNode[],
): string | undefined {
	if (getParams.length !== setterIndexParams.length) {
		return `Expected ${pluralizeCount(getParams.length, 'index parameter')}, but found ${setterIndexParams.length}.`;
	}
	for (let i = 0; i < getParams.length; i++) {
		const expected = getParams[i];
		const actual = setterIndexParams[i];
		if (!expected || !actual) {
			continue;
		}
		if (expected.isArray !== actual.isArray) {
			return `Index parameter ${i + 1} array shape must match.`;
		}
		if (effectivePassingMode(expected) !== effectivePassingMode(actual)) {
			return `Index parameter ${i + 1} passing mode must match.`;
		}
		const typeReason = propertyParameterTypeMismatch(expected, actual, i + 1);
		if (typeReason) {
			return typeReason;
		}
		const nameReason = propertyParameterNameMismatch(expected, actual, i + 1);
		if (nameReason) {
			return nameReason;
		}
	}
	return undefined;
}

/**
 * An index parameter keeps its name across the property's procedures: `Get
 * P(ByVal i As Long)` with `Let P(ByVal k As Long, ...)` is "Definitions of
 * property procedures for the same property are inconsistent" (issue #266,
 * measured in Excel 16.0). Case does not count, and the value parameter may
 * have any name.
 */
function propertyParameterNameMismatch(expected: ParameterNode, actual: ParameterNode, index: number): string | undefined {
	const bare = (name: string): string => name.replace(/^\[|\]$/g, '').toLowerCase();
	return bare(expected.name) === bare(actual.name)
		? undefined
		: `Index parameter ${index} must keep its name: expected '${expected.name}', found '${actual.name}'.`;
}

function propertyParameterTypeMismatch(
	expected: ParameterNode,
	actual: ParameterNode,
	index: number,
): string | undefined {
	const expectedType = normalizeType(expected.asType) ?? 'variant';
	const actualType = normalizeType(actual.asType) ?? 'variant';
	if (expectedType === actualType) {
		return undefined;
	}
	const scalarOrVariant =
		(expectedType === 'variant' || isKnownScalarType(expectedType)) &&
		(actualType === 'variant' || isKnownScalarType(actualType));
	if (!scalarOrVariant) {
		return undefined;
	}
	return `Index parameter ${index} type must match: expected ${expected.asType ?? 'Variant'}, found ${actual.asType ?? 'Variant'}.`;
}

function effectivePassingMode(param: ParameterNode): 'byval' | 'byref' {
	return param.byVal ? 'byval' : 'byref';
}

function propertyProcedureLabel(kind: ProcedureNode['procKind']): string {
	switch (kind) {
		case 'PropertyGet':
			return 'Property Get';
		case 'PropertyLet':
			return 'Property Let';
		case 'PropertySet':
			return 'Property Set';
		default:
			return 'Property';
	}
}

function procedureNameHit(source: string, proc: ProcedureNode): NameTokenHit | undefined {
	const header = firstLineSpan(source, proc.span);
	const toks = statementTokens(source, header);
	let i = 0;
	while (i < toks.length && PROC_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	const head = tokenText(toks[i]);
	if (head === 'property') {
		i += 2;
	} else if (head === 'sub' || head === 'function') {
		i++;
	}
	const tok = toks[i];
	const name = tok ? tokenName(tok) : undefined;
	return tok && name ? nameTokenHit(header, tok, name) : undefined;
}

function typeOrEnumNameHit(
	source: string,
	span: Span,
	keyword: 'type' | 'enum',
): NameTokenHit | undefined {
	const header = firstLineSpan(source, span);
	const toks = statementTokens(source, header);
	let i = 0;
	if (tokenText(toks[i]) === 'public' || tokenText(toks[i]) === 'private') {
		i++;
	}
	if (tokenText(toks[i]) === keyword) {
		i++;
	}
	const tok = toks[i];
	const name = tok ? tokenName(tok) : undefined;
	return tok && name ? nameTokenHit(header, tok, name) : undefined;
}

function declareNameHit(source: string, span: Span): NameTokenHit | undefined {
	const toks = statementTokens(source, span);
	const kindIndex = toks.findIndex(
		(tok) => tokenText(tok) === 'sub' || tokenText(tok) === 'function',
	);
	const tok = kindIndex >= 0 ? toks[kindIndex + 1] : undefined;
	const name = tok ? tokenName(tok) : undefined;
	return tok && name ? nameTokenHit(span, tok, name) : undefined;
}

function firstLineSpan(source: string, span: Span): Span {
	const nl = source.indexOf('\n', span.start);
	return {
		start: span.start,
		end: nl === -1 ? span.end : Math.min(nl, span.end),
	};
}

/**
 * Rule: a variable declaration cannot include an inline initializer. VBA has no
 * VB.NET-style `Dim x As Long = 1`; the `= value` is a syntax error. `Const`
 * legitimately uses `=` and is skipped. Detection walks every non-Const
 * VariableGroup (module level and inside procedure bodies) and looks for a
 * top-level `=` operator in the group's source slice - a declaration list has no
 * other lawful place for a depth-0 `=`.
 */
export function checkDimInitializer(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspect = (group: VariableGroupNode): void => {
		if (group.isConst) {
			return; // Const requires `=`; not an error.
		}
		const at = topLevelAssignOffset(source, group.span);
		if (at !== undefined) {
			push(
				'dimInitializer',
				'A variable declaration cannot include an initializer in VBA; assign the value in a separate statement.',
				{ start: at, end: at + 1 },
			);
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspect(member);
		} else if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspect, activity);
		}
	}
}

/**
 * Rule: once a declaration's `As <type>` clause is complete, another token in
 * the same logical statement must be introduced by real declaration syntax
 * (`=`, `,`, `:`/newline, etc.). A bare identifier after a complete type name,
 * as in `Dim s As String junk`, is VBE Compile `Syntax error`.
 *
 * This rule is intentionally narrow. It validates the token shape around the
 * `As` clause only; broad unknown type-name resolution belongs to the
 * project-wide binder. Recognized fixed-length String suffixes are consumed by
 * the shared suffix parser before trailing-token detection; their literal size
 * bounds are checked by `checkFixedLengthStringBounds`.
 */
export function checkUnexpectedDeclarationTokens(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspect = (span: Span, allowEquals: boolean): void => {
		const hit = unexpectedTokenAfterDeclarationType(source, span, allowEquals);
		if (!hit) {
			return;
		}
		push(
			'unexpectedDeclarationToken',
			`Unexpected token '${hit.text}' after a complete declaration type; this will fail to compile as a syntax error.`,
			hit.span,
		);
	};

	// A declaration that is not `name [As type]` (issue #234, measured in
	// Excel 16.0): "Syntax error" in a procedure, "Expected: end of statement"
	// at module level.
	const inspectGroup = (group: VariableGroupNode, error = 'Syntax error'): void => {
		for (const decl of group.declarations) {
			inspect(decl.span, true);
			const junk = declarationJunk(source, decl.span, group.isConst === true);
			if (junk) {
				push(
					'unexpectedDeclarationToken',
					`Unexpected '${junk.text}' after '${decl.name}': ${junk.why}. This is a VBE compile error: ${junk.error ?? error}.`,
					junk.span,
				);
			}
		}
	};

	const firstProcedure = mod.members.find((m) => m.kind === 'Procedure');
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			// After a procedure the VBE says only "Syntax error".
			inspectGroup(member, firstProcedure && member.span.start > firstProcedure.span.start ? 'Syntax error' : 'Expected: end of statement');
			continue;
		}
		if (member.kind === 'Type') {
			for (const field of member.fields) {
				inspectTypeField(field, inspect);
			}
			continue;
		}
		if (member.kind === 'Procedure') {
			for (const param of member.params) {
				inspectParameter(source, param, inspect);
			}
			forEachVariableGroup(member.body, inspectGroup, activity);
		}
	}
}

type TypeDeclarationSuffixNode =
	| ParameterNode
	| ProcedureNode
	| TypeFieldNode
	| VariableDeclNode;

/**
 * Rule: several declaration forms reject a legacy type-declaration character on
 * the name (`name$`, `count&`, etc.) when the same declaration also has an
 * explicit `As` clause. Property Get declarations are VBE-verified controls and
 * stay quiet.
 */
export function checkTypeDeclarationCharacterAsClause(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (node: TypeDeclarationSuffixNode, label: string): void => {
		if (!node.typeSuffix || !node.hasAsClause) {
			return;
		}
		push(
			'typeDeclarationCharacterAsClause',
			`${label} '${node.name}' combines type-declaration character '${node.typeSuffix}' with an As clause; use only one type declaration form.`,
			node.typeSuffixSpan ?? node.span,
		);
	};

	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			report(decl, group.isConst ? 'Const declaration' : 'Declaration');
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			for (const field of member.fields) {
				report(field, 'Type field');
			}
			continue;
		}
		if (member.kind === 'Procedure') {
			if (member.procKind === 'Function') {
				report(member, 'Function');
			}
			for (const param of member.params) {
				report(param, 'Parameter');
			}
			forEachVariableGroup(member.body, inspectGroup, activity);
		}
	}
}

const FIXED_LENGTH_STRING_MIN = 1;

const FIXED_LENGTH_STRING_MAX = 65526;

/**
 * Rule: fixed-length String sizes must be in VBE's accepted range when the
 * length is a decimal literal or a same-procedure/module Const/Enum member
 * whose value can be reduced to a deterministic integer expression. Broader
 * constant-expression semantics remain deferred.
 */
export function checkFixedLengthStringBounds(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	procedureFilter?: (member: ProcedureNode) => boolean,
): void {
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity);
	const inspectDeclaration = (
		decl: VariableDeclNode | TypeFieldNode,
		constants: ReadonlyMap<string, number | undefined>,
	): void => {
		if (decl.fixedLength === undefined || isInactiveNode(activity, decl)) {
			return;
		}
		const value = resolveFixedLengthStringSize(decl.fixedLength, constants);
		if (value === undefined) {
			return;
		}
		if (value >= FIXED_LENGTH_STRING_MIN && value <= FIXED_LENGTH_STRING_MAX) {
			return;
		}
		push(
			'fixedLengthStringSize',
			`Fixed-length String size must be between ${FIXED_LENGTH_STRING_MIN} and ${FIXED_LENGTH_STRING_MAX} characters; got ${value}.`,
			fixedLengthStringLengthSpan(source, decl.span) ?? decl.span,
		);
	};

	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			inspectDeclaration(decl, moduleConstants);
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && procedureFilter && !procedureFilter(member)) { continue; }
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			for (const field of member.fields) {
				inspectDeclaration(field, moduleConstants);
			}
			continue;
		}
		if (member.kind === 'Procedure') {
			const procedureConstants = new Map(moduleConstants);
			collectBodyLiteralIntegerConstants(member.body, procedureConstants, activity);
			forEachVariableGroup(
				member.body,
				(group) => {
					for (const decl of group.declarations) {
						inspectDeclaration(decl, procedureConstants);
					}
				},
				activity,
			);
		}
	}
}

function fixedLengthStringLengthSpan(source: string, span: Span): Span | undefined {
	const toks = statementTokens(source, span);
	const asIndex = toks.findIndex((t) => tokenText(t) === 'as');
	if (asIndex < 0) {
		return undefined;
	}
	let typeStart = asIndex + 1;
	if (tokenText(toks[typeStart]) === 'new') {
		typeStart++;
	}
	const fixed = parseFixedLengthStringType(toks, typeStart);
	const token = fixed ? toks[fixed.lengthIndex] : undefined;
	return token ? absoluteSpan(span, token) : undefined;
}

/**
 * What stands after a declared name where the VBE takes nothing: a second
 * word, `Dim asdf qwer`; no type name after As, `Private v As 123`; or a
 * second value in a Const, `Const K = asdf qwer` (issue #234). A complete
 * type followed by more is unexpectedTokenAfterDeclarationType's.
 */
function declarationJunk(
	source: string,
	span: Span,
	isConst: boolean,
): { text: string; span: Span; why: string; error?: string } | undefined {
	const toks = statementTokens(source, span);
	let i = tokenText(toks[0]) === 'withevents' ? 1 : 0;
	const name = toks[i];
	// A name that is no identifier, or runs on into what follows (`_name`,
	// `1value`, `user-name`), is the identifier rules' to report.
	if (!name || !isDeclarationTypeNameToken(name)) {
		return undefined;
	}
	i++;
	if (toks[i] && toks[i].start === name.end && /^[$%&!#@]$/.test(toks[i].rawText)) {
		i++;
	} else if (toks[i] && toks[i].start === name.end && toks[i].rawText !== '(') {
		return undefined;
	}
	if (toks[i]?.rawText === '(') {
		const close = matchParenFrom(toks, i);
		if (close < 0) {
			return undefined;
		}
		i = close + 1;
	}
	const next = toks[i];
	if (!next) {
		return undefined;
	}
	if (tokenText(next) === 'as') {
		const type = toks[tokenText(toks[i + 1]) === 'new' ? i + 2 : i + 1];
		// `As (Long)` is "Syntax error" (issue #236).
		return type && !isDeclarationTypeNameToken(type)
			? { text: type.rawText, span: absoluteSpan(span, type), why: 'As needs a type name', error: type.rawText === '(' ? 'Syntax error' : 'Expected: New or type name' }
			: undefined;
	}
	if (next.rawText === '=') {
		const at = isConst ? juxtaposedValueIndex(toks, i + 1) : -1;
		return at < 0 ? undefined : { text: toks[at].rawText, span: absoluteSpan(span, toks[at]), why: 'a Const takes one value', error: 'Expected: end of statement' };
	}
	return { text: next.rawText, span: absoluteSpan(span, next), why: 'a declaration takes As and a type there, or nothing' };
}

function inspectTypeField(
	field: TypeFieldNode,
	inspect: (span: Span, allowEquals: boolean) => void,
): void {
	// Scan after the name: a member may be named As (`As As Long`).
	inspect({ start: field.nameSpan?.end ?? field.span.start, end: field.span.end }, false);
}

function inspectParameter(
	source: string,
	param: ParameterNode,
	inspect: (span: Span, allowEquals: boolean) => void,
): void {
	if (parameterArrayAsTypeSyntaxHit(source, param)) {
		return;
	}
	inspect(param.span, true);
}

function unexpectedTokenAfterDeclarationType(
	source: string,
	span: Span,
	allowEquals: boolean,
): { text: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	const asIndex = toks.findIndex((t) => tokenText(t) === 'as');
	if (asIndex < 0) {
		return undefined;
	}

	let i = asIndex + 1;
	if (tokenText(toks[i]) === 'new') {
		i++;
	}

	const typeStart = i;
	i = consumeDeclarationTypeName(toks, i);
	if (i === typeStart) {
		return undefined;
	}

	const fixedLengthString = parseFixedLengthStringType(toks, typeStart);
	if (fixedLengthString && fixedLengthString.endIndex > i) {
		i = fixedLengthString.endIndex;
	}

	const next = toks[i];
	if (!next) {
		return undefined;
	}
	if (allowEquals && next.kind === 'operator' && next.rawText === '=') {
		return undefined;
	}

	return {
		text: next.rawText,
		span: absoluteSpan(span, next),
	};
}

function consumeDeclarationTypeName(toks: VbaToken[], start: number): number {
	if (!isDeclarationTypeNameToken(toks[start])) {
		return start;
	}
	let i = start + 1;
	for (;;) {
		if (toks[i]?.rawText !== '.') {
			return i;
		}
		if (!isDeclarationTypeNameToken(toks[i + 1])) {
			return start;
		}
		i += 2;
	}
}

function isDeclarationTypeNameToken(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	return (
		tok.kind === 'identifier' ||
		tok.kind === 'keyword' ||
		tok.kind === 'bracketedIdentifier'
	);
}

/** The Scripting Runtime's types, which missing-library-reference judges (issue #349). */
const SCRIPTING_TYPE_NAMES: ReadonlySet<string> = new Set(['dictionary', 'filesystemobject', 'textstream']);

export function checkInvalidAsTypeNames(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	opts: AnalyzeModuleOptions,
	push: PushFn,
): void {
	const withEventsNewDeclarationSpans = collectWithEventsNewDeclarationSpans(mod, activity);
	let variables: Set<string> | undefined;
	let ownTypes: Set<string> | undefined;
	let qualifiedEnumTypes: Set<string> | undefined;
	let libraries: readonly (ReadonlySet<string> | undefined)[] | undefined;
	for (const ref of collectTypeNameReferences(source)) {
		if (activity?.isInactive(ref.span)) {
			continue;
		}
		const lookupName = typeReferenceLookupName(ref);
		if (ref.kind === 'declaration' && ref.qualifier && !qualifiedEnumTypes) {
			qualifiedEnumTypes = new Set((opts.projectTypes ?? []).filter(type => type.kind === 'enum' && type.moduleName)
				.map(type => `${type.moduleName}.${type.name}`.toLowerCase()));
			for (const member of activeModuleMembers(mod,activity)) {
				if (member.kind === 'Enum') { qualifiedEnumTypes.add(`${opts.moduleName ?? 'Module'}.${member.name}`.toLowerCase()); }
			}
		}
		const qualifiedSourceEnum = ref.kind === 'declaration' && ref.qualifier && qualifiedEnumTypes?.has(lookupName.toLowerCase());
		if (qualifiedSourceEnum && !resolveTypeName(lookupName, {model:opts.hostModel})) {
			push('invalidAsTypeName', `'${lookupName}' qualifies a source Enum with its module name. Use '${ref.name}' as the type name. This is a VBE compile error: User-defined type not defined.`, ref.span);
			continue;
		}
		const resolved = resolveTypeName(lookupName, {
			projectTypes: opts.projectTypes,
			model: opts.hostModel,
		});
		if (resolved?.kind === 'ambiguous') {
			push(
				'invalidAsTypeName',
				`'${ref.name}' is ambiguous because multiple visible project types use that name.`,
				ref.span,
			);
			continue;
		}
		if (
			resolved &&
			isNewTypeReference(ref.kind) &&
			!isCreatableTypeCompletion(resolved) &&
			resolved.kind !== 'host'
		) {
			if (
				ref.kind === 'newDeclaration' &&
				withEventsNewDeclarationSpans.some((span) => containsSpan(span, ref.span))
			) {
				continue;
			}
			push(
				'invalidNewTypeName',
				`'${ref.name}' is ${typeKindLabelForNew(resolved.kind)} and cannot be used with New. New can create project classes and UserForms only.`,
				ref.span,
			);
			continue;
		}
		if (resolved) {
			continue;
		}
		// A Private Type or Enum of another module, bare or qualified, and any
		// name qualified by a variable, are no type here (issue #490, measured
		// in Excel 16.0).
		if (opts.hiddenTypeNames?.has(lookupName.toLowerCase())) {
			push('invalidAsTypeName', `'${lookupName}' is Private to the module that declares it, so this module cannot use it as a type. This is a VBE compile error: User-defined type not defined.`, ref.span);
			continue;
		}
		if (ref.qualifier && (variables ??= declaredVariableNames(mod, activity)).has(ref.qualifier.toLowerCase())) {
			push('invalidAsTypeName', `'${ref.qualifier}' is a variable, and a variable never qualifies a type. This is a VBE compile error: User-defined type not defined.`, ref.span);
			continue;
		}
		if (isReservedIdentifier(ref.name)) {
			push(
				'invalidAsTypeName',
				`'${ref.name}' is a reserved VBA identifier, not a valid type name.`,
				ref.span,
			);
			continue;
		}
		if (resolveRuntimeFunction(ref.name)) {
			push(
				'invalidAsTypeName',
				`'${ref.name}' is a VBA runtime function, not a valid type name.`,
				ref.span,
			);
			continue;
		}
		if (opts.knownNonTypeNames?.has(ref.name.toLowerCase())) {
			push(
				'invalidAsTypeName',
				`'${ref.name}' resolves to a project declaration, but that declaration is not a type.`,
				ref.span,
			);
			continue;
		}
		// No type of the project and none of a referenced library spells it,
		// where every library the project references is one whose names are
		// all known (issue #234, measured in Excel 16.0).
		// The Scripting Runtime's own types are missing-library-reference's, which
		// names the reference to add.
		// The dir stream omits the implicit VBA and container-host libraries.
		// They still contribute types even when only stdole/Office are recorded.
		libraries ??= opts.referencedLibraries === undefined ? undefined :
			[...new Set(opts.referencedLibraries.length
				? ['VBA', opts.hostModel?.hostName ?? opts.host ?? 'Excel', ...opts.referencedLibraries] : opts.referencedLibraries)]
				.map((library) => libraryTypeNames(library));
		ownTypes ??= new Set(activeModuleMembers(mod, activity).filter((member) => member.kind === 'Type' || member.kind === 'Enum').map((member) => member.name.toLowerCase()));
		if (!ref.qualifier && !SCRIPTING_TYPE_NAMES.has(ref.name.toLowerCase()) && !ownTypes.has(ref.name.toLowerCase()) && libraries !== undefined && libraries.length > 0 && libraries.every((names) => names !== undefined && !names.has(ref.name.toLowerCase()))) {
			push(
				'invalidAsTypeName',
				`No type of this project and none of the libraries it references is named '${ref.name}'. This is a VBE compile error: User-defined type not defined.`,
				ref.span,
			);
		}
	}
}

/** The variables the module declares, at module level or in a procedure, lowercased. */
function declaredVariableNames(mod: ModuleNode, activity: ConditionalActivityTracker | undefined): Set<string> {
	const out = new Set<string>();
	const add = (group: VariableGroupNode): void => {
		if (!group.isConst) {
			for (const decl of group.declarations) {
				out.add(decl.name.toLowerCase());
			}
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			add(member);
		} else if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, add, activity);
		}
	}
	return out;
}

function collectWithEventsNewDeclarationSpans(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): Span[] {
	const spans: Span[] = [];
	const inspect = (group: VariableGroupNode): void => {
		if (!group.withEvents || isInactiveNode(activity, group)) {
			return;
		}
		for (const decl of group.declarations) {
			if (decl.isNew) {
				spans.push(decl.span);
			}
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspect(member);
			continue;
		}
		if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspect, activity);
		}
	}
	return spans;
}

function containsSpan(container: Span, inner: Span): boolean {
	return inner.start >= container.start && inner.end <= container.end;
}

function keywordSpan(source: string, span: Span, ...keywords: string[]): Span {
	const expected = new Set(keywords);
	const tok = statementTokensAfterLeadingLabel(source, span)
		.find((token) => expected.has(tokenText(token)));
	return tok ? absoluteSpan(span, tok) : firstTokenSpan(source, span);
}

/**
 * Returns the absolute offset of the first top-level `=` operator in the source
 * slice for `span`, or undefined. Parenthesised regions (array bounds, default
 * sub-expressions) are skipped so only a declaration-level `=` is reported.
 */
function topLevelAssignOffset(source: string, span: Span): number | undefined {
	const toks = statementTokens(source, span);
	let depth = 0;
	for (const t of toks) {
		const r = t.rawText;
		if (r === '(') {
			depth++;
		} else if (r === ')') {
			depth--;
		} else if (depth === 0 && t.kind === 'operator' && r === '=') {
			return span.start + t.start;
		}
	}
	return undefined;
}

/**
 * Rule: parameter-list constraints. A required parameter may not follow an
 * `Optional` one, `ParamArray` must be the final parameter, `ParamArray` cannot
 * be combined with Optional parameters in the same list, and explicitly typed
 * `ParamArray` elements must be Variant. These are read straight off the parsed
 * parameter flags, so they are deterministic.
 */
export function checkParameterOrder(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const params = member.params;
		const hasOptional = params.some((p) => p.optional);
		// The final parameter of a Property Let/Set is the assigned value: it is
		// mandatory by definition and exempt from the "required-after-optional"
		// constraint, so an Optional index parameter may legally precede it
		// (MS-VBAL 5.3.1.5). Only the index parameters obey the ordering rule.
		const lastIsValueParameter =
			member.procKind === 'PropertyLet' || member.procKind === 'PropertySet';
		let optionalSeen = false;
		for (let i = 0; i < params.length; i++) {
			const p = params[i];
			const arrayAsType = parameterArrayAsTypeSyntaxHit(source, p);
			if (arrayAsType) {
				push(
					'parameterArrayAsTypeSyntax',
					`Array parameter '${p.name}' must place parentheses after the parameter name, before the As clause; use '${p.name}() As ${arrayAsType.typeName}'.`,
					arrayAsType.span,
				);
				if (p.optional) {
					optionalSeen = true;
				}
				continue;
			}
			if (p.paramArray) {
				if (p.asType && normalizeType(p.asType) !== 'variant') {
					push(
						'paramArrayNonVariant',
						`ParamArray '${p.name}' elements must be Variant, but this parameter is declared As ${p.asType}.`,
						declaredNameSpan(source, p.span, p.name),
					);
				}
				if (hasOptional) {
					push(
						'paramArrayWithOptional',
						`ParamArray '${p.name}' cannot be used in the same parameter list as Optional arguments.`,
						declaredNameSpan(source, p.span, p.name),
					);
				}
				// A Property Let or Set takes its value after the ParamArray:
				// `Property Let P(ParamArray v() As Variant, ByVal x As Long)`
				// compiles (issue #266, measured in Excel 16.0).
				if (i !== params.length - 1 && !(lastIsValueParameter && i === params.length - 2)) {
					push(
						'paramArrayNotLast',
						`ParamArray '${p.name}' must be the last parameter.`,
						declaredNameSpan(source, p.span, p.name),
					);
				}
				continue;
			}
			if (p.optional) {
				optionalSeen = true;
				continue;
			}
			if (optionalSeen && !(lastIsValueParameter && i === params.length - 1)) {
				push(
					'requiredParamAfterOptional',
					`Parameter '${p.name}' must be Optional because it follows an Optional parameter.`,
					declaredNameSpan(source, p.span, p.name),
				);
			}
		}
	}
}

function parameterArrayAsTypeSyntaxHit(
	source: string,
	param: ParameterNode,
): { span: Span; typeName: string } | undefined {
	const toks = statementTokens(source, param.span);
	const asIndex = toks.findIndex((t) => tokenText(t) === 'as');
	if (asIndex < 0) {
		return undefined;
	}
	let typeStart = asIndex + 1;
	if (tokenText(toks[typeStart]) === 'new') {
		typeStart++;
	}
	const typeEnd = consumeDeclarationTypeName(toks, typeStart);
	if (typeEnd === typeStart) {
		return undefined;
	}
	const open = toks[typeEnd];
	const close = toks[typeEnd + 1];
	if (!open || !close || open.rawText !== '(' || close.rawText !== ')') {
		return undefined;
	}
	return {
		span: {
			start: param.span.start + open.start,
			end: param.span.start + close.end,
		},
		typeName: source.slice(param.span.start + toks[typeStart].start, param.span.start + toks[typeEnd - 1].end),
	};
}

/**
 * Rule: Optional parameter defaults must be compile-time compatible with their
 * declared type when the default expression is deterministic. VBE oracle
 * evidence rejects nonnumeric string defaults for numeric and Boolean
 * parameters as compile-time Type mismatch, while numeric strings remain valid.
 * Array parameters cannot be initialized from scalar defaults, and object
 * parameters default only to Nothing when the object type is known.
 */
export function checkParameterDefaultValues(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const objectType = createObjectAssignmentTypeResolver(memberCtx);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		for (const param of member.params) {
			if (!param.defaultRaw || !param.asType) {
				continue;
			}
			const defaultTokens = parameterDefaultTokens(source, param);
			if (!defaultTokens) {
				continue;
			}
			const actual = inferArgumentType(defaultTokens.tokens, param.span.start, new Map(), new Map());
			if (!actual) {
				continue;
			}
			const reason = parameterDefaultIncompatibilityReason(param, actual, objectType);
			if (!reason) {
				continue;
			}
			push(
				'parameterDefaultTypeMismatch',
				`Optional parameter '${param.name}' expects ${parameterDefaultExpectedLabel(param)}, but its default value is ${actual.label}. ${reason}`,
				defaultTokens.span,
			);
		}
	}
}

/**
 * Optional parameter defaults must be constant expressions (MS-VBAL 5.3.1.5 /
 * VBE "Constant expression required"). Flags a default that is provably
 * non-constant - a function/array call (`name(...)`), `New`, or `AddressOf`.
 * Bare identifiers and member references (`Module.CONST`, `MyEnum.Value`) are
 * left alone because they may be constants, so this stays no-false-positive.
 * Object-typed parameters are skipped: their defaults are owned by the
 * `parameter-default-type-mismatch` rule ("must be Nothing"), avoiding a
 * double diagnostic.
 */
export function checkNonConstantParameterDefaults(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const objectType = createObjectAssignmentTypeResolver(memberCtx);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		for (const param of member.params) {
			if (!param.defaultRaw) {
				continue;
			}
			if (objectType(param.asType)) {
				continue;
			}
			const defaultTokens = parameterDefaultTokens(source, param);
			if (!defaultTokens) {
				continue;
			}
			const nonConstant = nonConstantDefaultElement(defaultTokens.tokens, param.span.start, 'enumOrOptional');
			if (!nonConstant) {
				continue;
			}
			push(
				'parameterDefaultNotConstant',
				`Optional parameter '${param.name}' default must be a constant expression; ${nonConstant.label} is not constant.`,
				nonConstant.span,
			);
		}
	}
}

/**
 * The value of a Const declaration must be a constant expression (MS-VBAL 5.2.4
 * / VBE "Constant expression required"). Flags a Const value that is provably
 * non-constant - a function/array call (`name(...)`), `New`, or `AddressOf` -
 * at module level and procedure-local (including nested blocks). Bare and
 * qualified identifiers (`OTHER_CONST`, `Module.CONST`, `MyEnum.Value`) are
 * left alone because they may reference constants, so this stays
 * no-false-positive. Literals, string concatenation, and arithmetic/grouping
 * are constant expressions and never flagged. VBA's functions that take no
 * argument, `Now`, `Date`, `Time`, `Timer` and `Rnd`, are calls without the
 * parentheses unless the module declares the name (issue #255, measured).
 */
export function checkNonConstantConstValues(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const declared = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			member.declarations.forEach((decl) => declared.add(decl.name.toLowerCase()));
		} else if ('name' in member && typeof member.name === 'string') {
			declared.add(member.name.toLowerCase());
		}
	}
	const inspectGroup = (group: VariableGroupNode): void => {
		if (!group.isConst) {
			return;
		}
		for (const decl of group.declarations) {
			if (decl.defaultRaw === undefined || isInactiveNode(activity, decl)) {
				continue;
			}
			const valueTokens = valueTokensAfterEquals(source, decl.span);
			if (!valueTokens) {
				continue;
			}
			const nonConstant = nonConstantDefaultElement(valueTokens.tokens, decl.span.start, 'const')
				?? argumentlessFunction(valueTokens.tokens, decl.span.start, declared);
			if (!nonConstant) {
				continue;
			}
			push(
				'constValueNotConstant',
				`Const '${decl.name}' value must be a constant expression; ${nonConstant.label} is not constant.`,
				nonConstant.span,
			);
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
			continue;
		}
		if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspectGroup, activity);
		}
	}
}

/** VBA's functions measured as refused in a Const without parentheses (issue #255). */
const ARGUMENTLESS_FUNCTIONS: ReadonlySet<string> = new Set(['now', 'date', 'time', 'timer', 'rnd']);

/** `Const K = Now`: a VBA function named without parentheses, which still calls it. */
function argumentlessFunction(
	toks: readonly VbaToken[],
	base: number,
	declared: ReadonlySet<string>,
): { label: string; span: Span } | undefined {
	for (let i = 0; i < toks.length; i++) {
		const word = tokenText(toks[i]);
		if (ARGUMENTLESS_FUNCTIONS.has(word) && !declared.has(word) && toks[i - 1]?.rawText !== '.') {
			return { label: `'${toks[i].rawText}', a VBA function evaluated as the code runs,`, span: { start: base + toks[i].start, end: base + toks[i].end } };
		}
	}
	return undefined;
}

/**
 * Enum member values must be constant expressions (MS-VBAL 5.2.3.4 / VBE
 * "Constant expression required"). Flags a member initializer that is provably
 * non-constant - a function/array call (`name(...)`), `New`, or `AddressOf`.
 * Bare and qualified identifiers stay quiet (they may be constants). Implicit
 * members (no `=`) are auto-numbered and never checked. No-false-positive by the
 * same gating as the Optional-default and Const rules.
 */
export function checkNonConstantEnumMemberValues(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let stringConsts: Map<string, string | undefined> | undefined;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Enum') {
			continue;
		}
		stringConsts ??= moduleStringConstants(source, mod, activity);
		for (const enumMember of member.members) {
			// Skip members in an inactive #If branch (the parser models Enum-body
			// directives, so the activity tracker resolves these by offset).
			if (enumMember.valueRaw === undefined || isInactiveNode(activity, enumMember)) {
				continue;
			}
			const valueTokens = valueTokensAfterEquals(source, enumMember.span);
			if (!valueTokens) {
				continue;
			}
			const nonConstant = nonConstantDefaultElement(valueTokens.tokens, enumMember.span.start, 'enumOrOptional');
			if (nonConstant) {
				push(
					'enumMemberNotConstant',
					`Enum member '${enumMember.name}' value must be a constant expression; ${nonConstant.label} is not constant.`,
					nonConstant.span,
				);
				continue;
			}
			const text = constantStringValue(valueTokens.tokens, stringConsts);
			if (text !== undefined && stringIsNeverNumeric(text)) {
				push(
					'enumMemberTypeMismatch',
					`Enum member '${enumMember.name}' is the string "${text}", and an Enum member is a Long. This is a VBE compile error: Type mismatch.`,
					valueTokens.span,
				);
			}
		}
	}
}

/**
 * The module's Consts whose value is a string: `Const S As String = "x"` or
 * one built from others with `&`. Keyed by lowercased name; a name declared
 * twice maps to undefined.
 */
function moduleStringConstants(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): Map<string, string | undefined> {
	const out = new Map<string, string | undefined>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'VariableGroup' || !member.isConst) {
			continue;
		}
		for (const decl of member.declarations) {
			const tokens = decl.defaultRaw === undefined ? undefined : valueTokensAfterEquals(source, decl.span)?.tokens;
			const key = decl.name.toLowerCase();
			out.set(key, out.has(key) || !tokens ? undefined : constantStringValue(tokens, out));
		}
	}
	return out;
}

/**
 * A constant expression's value when it is a string: a literal, a Const that
 * is one, or those joined with `&`. Undefined for anything else.
 */
function constantStringValue(
	tokens: readonly VbaToken[],
	stringConsts: ReadonlyMap<string, string | undefined> | undefined,
): string | undefined {
	const parts = tokens.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
	let out = '';
	for (let i = 0; i < parts.length; i++) {
		const tok = parts[i];
		if (i % 2 === 1) {
			if (tok.rawText !== '&') {
				return undefined;
			}
			continue;
		}
		if (tok.kind === 'stringLiteral') {
			out += tok.rawText.slice(1, -1).replace(/""/g, '"');
		} else if (tok.kind === 'identifier') {
			const value = stringConsts?.get(tok.rawText.toLowerCase());
			if (value === undefined) {
				return undefined;
			}
			out += value;
		} else {
			return undefined;
		}
	}
	return parts.length % 2 === 1 ? out : undefined;
}

/**
 * Whether no locale could read the string as a number: it has no digit and is
 * not a `&H`/`&O` literal. Measured in Excel 16.0 (issue #210): `"1"` and
 * `"&H10"` are Longs to VBA, and `"x"`, `""` and `"True"` are a Type mismatch,
 * though CLng("True") runs.
 */
function stringIsNeverNumeric(text: string): boolean {
	const trimmed = text.trim();
	return !/\d/.test(trimmed) && !trimmed.startsWith('&');
}

/**
 * Operator keywords (And, Or, Not, Mod, Xor, Eqv, Imp, Is, Like, TypeOf, plus
 * New/AddressOf) lex as `keyword` but are never callable names. They must be
 * excluded from the call heuristic below, otherwise a legal constant expression
 * like `6 And (3)` reads as a bogus call `And(...)`. New/AddressOf are still
 * flagged by the dedicated branch above this set's use.
 */
const OPERATOR_KEYWORD_WORDS = new Set(OPERATOR_IDENTIFIERS.map((word) => word.toLowerCase()));

/**
 * The intrinsic functions the VBE folds inside an Enum member value and an
 * Optional parameter default, where it refuses every call in a Const. Measured
 * one by one in Excel 16.0 (build 20326, 2026-09-26; issue #112): these
 * sixteen compile in both positions, while Asc, AscW, Chr, Val, Sqr, RGB,
 * Round, IIf, Hex, Oct, InStr, StrComp, CDec, DateSerial, Choose, Mid, Left,
 * UCase, Str, Trim, Format, Replace, String, Space, Now, Timer, Rnd and Array
 * are "Constant expression required" there too.
 */
const CONSTANT_FOLDED_INTRINSICS: ReadonlySet<string> = new Set([
	'len', 'lenb', 'abs', 'int', 'fix', 'sgn',
	'cint', 'clng', 'clnglng', 'cbyte', 'cbool', 'cdbl', 'csng', 'ccur', 'cvar', 'cdate',
]);

function nonConstantDefaultElement(
	tokens: VbaToken[],
	baseOffset: number,
	position: 'const' | 'enumOrOptional',
): { label: string; span: Span } | undefined {
	for (let i = 0; i < tokens.length; i++) {
		const tok = tokens[i];
		const word = (tok.canonicalText ?? tok.rawText).toLowerCase();
		if (tok.kind === 'keyword' && (word === 'new' || word === 'addressof')) {
			return {
				label: `'${tok.rawText}'`,
				span: { start: baseOffset + tok.start, end: baseOffset + tokens[tokens.length - 1].end },
			};
		}
		const isName =
			tok.kind === 'identifier' || tok.kind === 'keyword' || tok.kind === 'bracketedIdentifier';
		const isOperatorKeyword = tok.kind === 'keyword' && OPERATOR_KEYWORD_WORDS.has(word);
		if (
			position === 'enumOrOptional'
			&& CONSTANT_FOLDED_INTRINSICS.has(word)
			&& tokens[i - 1]?.rawText !== '.'
		) {
			continue;
		}
		if (isName && !isOperatorKeyword && tokens[i + 1]?.rawText === '(') {
			const closeIndex = matchParenFrom(tokens, i + 1);
			const endTok = closeIndex >= 0 ? tokens[closeIndex] : tokens[i + 1];
			return {
				label: `the call '${tok.rawText}(...)'`,
				span: { start: baseOffset + tok.start, end: baseOffset + endTok.end },
			};
		}
	}
	return undefined;
}

function parameterDefaultTokens(
	source: string,
	param: ParameterNode,
): { tokens: VbaToken[]; span: Span } | undefined {
	return valueTokensAfterEquals(source, param.span);
}

/**
 * Tokenizes the slice for `span`, finds the top-level `=`, and returns the
 * tokens after it (the value/default expression) plus their absolute span.
 * Shared by the Optional-default, Const, and Enum-member constant-expression
 * rules. Returns undefined when there is no top-level `=` or nothing follows it.
 */
function valueTokensAfterEquals(
	source: string,
	span: Span,
): { tokens: VbaToken[]; span: Span } | undefined {
	const toks = statementTokens(source, span);
	const eq = topLevelOperatorIndex(toks, '=');
	if (eq < 0 || eq + 1 >= toks.length) {
		return undefined;
	}
	const tokens = toks.slice(eq + 1);
	return {
		tokens,
		span: spanForTokens(tokens, span.start),
	};
}

function parameterDefaultIncompatibilityReason(
	param: ParameterNode,
	actual: InferredArgumentType,
	objectType: ReturnType<typeof createObjectAssignmentTypeResolver>,
): string | undefined {
	if (param.isArray && isKnownScalarDefaultType(actual.type)) {
		return 'Optional array parameter defaults cannot be scalar values.';
	}
	const expectedRaw = param.asType;
	if (!expectedRaw) {
		return undefined;
	}
	const expectedObject = objectType(expectedRaw);
	if (expectedObject) {
		return normalizeType(actual.type) === 'nothing'
			? undefined
			: 'Optional object parameter defaults must be Nothing.';
	}
	const reason = incompatibilityReason(expectedRaw, actual);
	if (!reason || !/string literal/i.test(actual.label)) {
		return undefined;
	}
	return 'This is a VBE compile error: Type mismatch.';
}

function parameterDefaultExpectedLabel(param: ParameterNode): string {
	const base = param.asType ?? 'Variant';
	return param.isArray ? `${base}()` : base;
}

function isKnownScalarDefaultType(type: string | undefined): boolean {
	const normalized = normalizeType(type);
	return !!normalized && isKnownScalarType(normalized);
}

function isNewTypeReference(kind: TypeNameReferenceKind): boolean {
	return kind === 'newExpression' || kind === 'newDeclaration';
}

function typeKindLabelForNew(kind: TypeCompletionKind): string {
	switch (kind) {
		case 'primitive':
			return 'a VBA primitive type';
		case 'external':
			return 'an external interface type';
		case 'host':
			return 'a host object-model type';
		case 'document':
			return 'a document module type';
		case 'enum':
			return 'an Enum type';
		case 'userType':
			return 'a user-defined Type';
		case 'ambiguous':
			return 'an ambiguous project type';
		case 'module':
			return 'a module qualifier';
		case 'class':
		case 'userform':
			return 'a creatable project type';
	}
}

/**
 * Rule: an `Option` statement may not follow a procedure.
 *
 * Only a procedure closes the window. Measured in Excel 16.0 (build 20326,
 * 2026-09-26): each of Option Explicit, Base, Compare and Private Module
 * compiles after a Const, a module variable, a Type, an Enum, a Declare and a
 * Deftype statement (issue #113 opened with `DefLng A-Z` above
 * `Option Explicit`), and is refused only after `End Sub` / `End Function`
 * with "Only comments may appear after End Sub, End Function, or End
 * Property". The rule used to treat every declaration as closing the window.
 */
export function checkOptionPlacement(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Procedures that precede the Option under test AND could be compiled
	// beside it: a procedure in the other arm of a chain closes no window,
	// because only one arm is ever built (issues/58).
	const proceduresAbove: Span[] = [];
	// `Option Base` alone has a second closer: a module-level array already
	// dimensioned above it ("Array already dimensioned", measured 2026-09-26).
	const arraysAbove: Span[] = [];
	const compiledWith = (priors: readonly Span[], member: Span): boolean =>
		priors.some((prior) => !activity?.mutuallyExclusive(prior, member));
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Option') {
			if (compiledWith(proceduresAbove, member.span)) {
				push(
					'optionAfterDeclaration',
					'Option statements must appear before the first procedure; only comments may follow End Sub, End Function, or End Property.',
					firstTokenSpan(source, member.span),
				);
			} else if (/^base\b/i.test(member.optionText.trim()) && compiledWith(arraysAbove, member.span)) {
				push(
					'optionAfterDeclaration',
					"'Option Base' must come before any array declaration: an array above it is already dimensioned.",
					firstTokenSpan(source, member.span),
				);
			}
			continue;
		}
		if (member.kind === 'Procedure') {
			proceduresAbove.push(member.span);
		} else if (member.kind === 'VariableGroup' && member.declarations.some((decl) => decl.isArray)) {
			arraysAbove.push(member.span);
		}
	}
}

/**
 * Rule: a user-defined Type must declare at least one member. VBE rejects an
 * empty Type with "User-defined type without members not allowed" (MS-VBAL
 * 5.2.3.3; oracle-verified `empty_type_block_compile`). A Type whose only fields
 * sit in an inactive `#If` branch is also empty at compile time; an unknown
 * branch keeps the field active, so the block stays quiet (no false positive).
 * Unclosed Type blocks are left to the missing-`End Type` parse diagnostic.
 */
export function checkEmptyType(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Type' || !member.closed) {
			continue;
		}
		if (member.fields.some((field) => !isInactiveNode(activity, field))) {
			continue;
		}
		push(
			'emptyType',
			`Type '${member.name}' must declare at least one member.`,
			member.nameSpan ?? member.span,
		);
	}
}

/**
 * Rule: a module may declare each Option only once. VBE rejects a repeated
 * Option statement with "Duplicate Option statement" (MS-VBAL 5.2.1;
 * oracle-verified `duplicate_option_explicit_compile`). Keyed by Option category
 * (Explicit / Compare / Base / Private), so two `Option Compare` collide even
 * with different arguments, while distinct Options never do.
 */
export function checkDuplicateOptions(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Two Options in different arms of one `#If` chain are alternatives, so
	// only the arm matters, not whether the branch can be decided. Skipping
	// every undecidable branch went blind to a real repeat inside one arm.
	const options = [...activeModuleMembers(mod, activity)].filter((m) => m.kind === 'Option');
	reportRepeatedKeys(options, activity, {
		keyOf: (member) => {
			if (activity?.isInactive(member.span)) {
				return undefined;
			}
			const category = optionCategory(member).toLowerCase();
			return category || undefined;
		},
		spanOf: (member) => member.span,
		report: (repeat) => push(
			'duplicateOption',
			`Duplicate Option statement; only one 'Option ${optionCategory(repeat)}' is allowed per module.`,
			firstTokenSpan(source, repeat.span),
		),
	});
}

/** The word after `Option`, which is what may appear at most once. */
function optionCategory(member: { optionText: string }): string {
	return member.optionText.trim().split(/\s+/)[0] ?? '';
}

/**
 * Rule: an Option statement names one of the four directives VBA has, takes the
 * argument that one takes, and ends there. The parser kept whatever followed
 * `Option` without reading it, so `Option Explicit()` analyzed clean while the
 * module would not compile (issue #74).
 *
 * Every message below is the live VBE's own wording, read off the compile
 * dialog for each malformed form (Excel oracle probes, 2026-09-13):
 *
 *   Option                   Expected: Base or Compare or Explicit or Private
 *   Option Nonsense          Expected: Base or Compare or Explicit or Private
 *   Option Explicit()        Expected: end of statement
 *   Option Explicit Foo      Expected: end of statement
 *   Option Base              Expected: 0 or 1
 *   Option Base 2            Expected: 0 or 1
 *   Option Base 1 Extra      Expected: end of statement
 *   Option Compare Sideways  Expected: Text or Binary
 *   Option Private           Expected: Module
 *
 * `Option Compare Database` is the one the host decides. Access writes it into
 * every module it creates, and Excel refuses it with the same "Text or Binary"
 * as any other unknown argument, so it is reported only where the project names
 * a host that is not Access. A file no project claims names no host and is left
 * alone, the way issue #73 left a loose module's kind alone.
 */
export function checkOptionStatementForm(
	source: string,
	mod: ModuleNode,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const host = opts.host?.toLowerCase();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Option') {
			continue;
		}
		// A trailing comment is not trailing junk, and a line continuation is
		// trivia the lexer already attached to the token that follows it.
		const toks = statementTokens(source, member.span);
		const report = (index: number, message: string): void => {
			const tok = toks[index];
			push(
				'invalidOptionStatement',
				message,
				tok ? absoluteSpan(member.span, tok) : firstTokenSpan(source, member.span),
			);
		};
		/** True when the statement ends where the directive's form says it should. */
		const endsHere = (index: number, form: string): boolean => {
			if (toks.length > index) {
				report(index, `'${form}' is complete here; VBA expects the statement to end.`);
				return false;
			}
			return true;
		};
		const argument = (index: number): string | undefined => (
			toks[index] ? tokenText(toks[index]) : undefined
		);
		// toks[0] is `Option` itself; the directive it names follows it.
		const directive = argument(1);
		if (directive === undefined) {
			report(0, "'Option' names no directive; VBA expects Base, Compare, Explicit or Private.");
			continue;
		}
		switch (directive) {
			case 'explicit':
				endsHere(2, 'Option Explicit');
				break;
			case 'base': {
				const arg = argument(2);
				if (arg !== '0' && arg !== '1') {
					report(arg === undefined ? 1 : 2, "'Option Base' takes 0 or 1.");
					break;
				}
				endsHere(3, 'Option Base');
				break;
			}
			case 'compare': {
				const arg = argument(2);
				if (arg === 'database') {
					if (host !== undefined && host !== 'access') {
						report(2, "'Option Compare Database' is an Access directive; this project's host takes Binary or Text.");
						break;
					}
				} else if (arg !== 'binary' && arg !== 'text') {
					report(arg === undefined ? 1 : 2, "'Option Compare' takes Binary or Text.");
					break;
				}
				endsHere(3, 'Option Compare');
				break;
			}
			case 'private':
				if (argument(2) === 'module' && isObjectModuleKind(opts.moduleKind)) {
					// Measured in Excel 16.0 (issue #124): "Option Private Module
					// not permitted in an object module".
					report(2, "'Option Private Module' is not permitted in a class, document or UserForm module.");
					break;
				}
				if (argument(2) !== 'module') {
					report(
						argument(2) === undefined ? 1 : 2,
						"'Option Private' is written 'Option Private Module'.",
					);
					break;
				}
				endsHere(3, 'Option Private Module');
				break;
			default:
				report(
					1,
					`'Option ${toks[1].canonicalText ?? toks[1].rawText}' is not an Option statement; `
					+ 'VBA expects Base, Compare, Explicit or Private.',
				);
				break;
		}
	}
}

/** VBA allows at most 60 parameters on a procedure. */
const MAX_PROCEDURE_PARAMETERS = 60;
/** In a class module the most is 59 (issue #210). */
const MAX_CLASS_PROCEDURE_PARAMETERS = 59;

/**
 * Rule: a procedure, Event or Declare may declare at most 60 parameters, and
 * at most 59 in a class module. VBE rejects one more with "Too many
 * arguments" (oracle-verified `corpus_arg_limit_001b_compile`; issue #210,
 * measured in Excel 16.0: a class Sub, Friend Sub, Function, Property Get,
 * Property Let with its value, Event and Private Declare each compile with
 * 59 and are refused with 60, and a ParamArray counts as one). Document
 * modules and UserForms could not be measured the same way, so they keep 60.
 */
export function checkTooManyParameters(
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const limit = moduleKind === 'class' ? MAX_CLASS_PROCEDURE_PARAMETERS : MAX_PROCEDURE_PARAMETERS;
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			(member.kind !== 'Procedure' && member.kind !== 'Event' && member.kind !== 'Declare')
			|| member.params.length <= limit
		) {
			continue;
		}
		const what = member.kind === 'Procedure' ? 'a procedure' : member.kind === 'Event' ? 'an Event' : 'a Declare';
		const subject = moduleKind === 'class' ? `In a class module, ${what}` : `${what[0].toUpperCase()}${what.slice(1)}`;
		push(
			'tooManyParameters',
			`${subject} may have at most ${limit} parameters; '${member.name}' declares ${member.params.length}.`,
			member.nameSpan ?? member.span,
		);
	}
}

/** VBA identifiers may be at most 255 characters. */
const MAX_IDENTIFIER_LENGTH = 255;

/**
 * Rule: a declared identifier may be at most 255 characters. VBE rejects a longer
 * name with "Identifier too long" (oracle-verified `corpus_name_limit_001b_compile`).
 * Pure length check over declared names; binder-independent, no false positives.
 */
export function checkIdentifierTooLong(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (name: string, span: Span): void => {
		if (name.length <= MAX_IDENTIFIER_LENGTH) {
			return;
		}
		push(
			'identifierTooLong',
			`Identifier '${name.slice(0, 24)}...' is ${name.length} characters; VBA allows at most ${MAX_IDENTIFIER_LENGTH}.`,
			span,
		);
	};
	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			report(decl.name, decl.nameSpan ?? decl.span);
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
		} else if (member.kind === 'Type') {
			report(member.name, member.nameSpan ?? member.span);
			for (const field of member.fields) {
				report(field.name, field.nameSpan ?? field.span);
			}
		} else if (member.kind === 'Enum') {
			report(member.name, member.nameSpan ?? member.span);
			for (const enumMember of member.members) {
				report(enumMember.name, enumMember.nameSpan ?? enumMember.span);
			}
		} else if (member.kind === 'Procedure') {
			report(member.name, member.nameSpan ?? member.span);
			for (const param of member.params) {
				report(param.name, param.nameSpan ?? param.span);
			}
			forEachVariableGroup(member.body, inspectGroup, activity);
		}
	}
}

/**
 * Rule family on user-defined-Type parameters, both oracle-verified:
 *  - `optional-udt-parameter`: an `Optional` parameter cannot be a UDT ("Invalid
 *    optional parameter"; `corpus_sig_007_compile`).
 *  - `byval-udt-parameter`: a non-optional `ByVal` parameter cannot be a UDT - a
 *    UDT must be passed `ByRef` ("User-defined type may not be passed ByVal";
 *    `corpus_api_vis_003_compile`). The oracle confirmed this is about `ByVal`,
 *    not type visibility: `ByRef` UDT parameters are accepted.
 * Both fire only when the parameter's declared type matches a Type declared in
 * this module (unambiguously a UDT), so they are no-false-positive; cross-module
 * type names are not resolved here and stay quiet. An `Optional ByVal` UDT param
 * reports only the Optional diagnostic (matching VBE). A Declare's parameters
 * are held to both (issue #253, measured in Excel 16.0: "User-defined type
 * may not be passed ByVal" and "Invalid optional parameter type").
 */
export function checkUdtParameterConstraints(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const udtNames = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Type') {
			udtNames.add(member.name.trim().toLowerCase());
		}
	}
	if (udtNames.size === 0) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' && member.kind !== 'Declare') {
			continue;
		}
		for (const param of member.params) {
			if (!param.asType || !udtNames.has(param.asType.trim().toLowerCase())) {
				continue;
			}
			if (param.optional) {
				push(
					'optionalUdtParameter',
					`Optional parameter '${param.name}' cannot be a user-defined type ('${param.asType}').`,
					param.nameSpan ?? param.span,
				);
			} else if (param.byVal) {
				push(
					'byvalUdtParameter',
					`User-defined type parameter '${param.name}' ('${param.asType}') cannot be passed ByVal; pass it ByRef.`,
					param.nameSpan ?? param.span,
				);
			}
		}
	}
}
