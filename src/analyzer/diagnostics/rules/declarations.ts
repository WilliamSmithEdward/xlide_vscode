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
import { tokenize } from '../../lexer/tokenize';
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
import {
	collectTypeNameReferences,
	type TypeNameReferenceKind,
	typeReferenceLookupName,
} from '../../semantic/typeSemanticTokens';
import type {
	AnalyzeModuleOptions,
	PushFn,
} from '../analysisContext';
import type { InferredArgumentType } from '../callExtraction';
import {
	collectBodyLiteralIntegerConstants,
	collectModuleLiteralIntegerConstants,
	resolveFixedLengthStringSize,
} from '../constExpr';
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
	resolveKnownObjectAssignmentType,
	spanForTokens,
} from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	declaredNameSpan,
	firstTokenSpan,
	forEachBodyStatement,
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
				forEachBodyStatement((node as { body: BodyNode[] }).body, inspectStatement, activity);
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
 * return type, Property Let value parameters must not be object references,
 * and Property Set value parameters must be object references.
 */
export function checkPropertySetterValueParameters(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	opts: AnalyzeModuleOptions,
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
		if (member.params.length > 0) {
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
			} else {
				const objectType = resolveKnownObjectAssignmentType(valueParam.asType, {
					projectClassMembers: opts.projectClassMembers,
					model: opts.hostModel,
				});
				if (objectType) {
					push(
						'propertyLetObjectValue',
						`Property Let '${member.name}' final value parameter '${valueParam.name}' must not be an object reference; use Property Set because it is declared As ${objectType.display}.`,
						declaredNameSpan(source, valueParam.span, valueParam.name),
					);
				}
			}
			continue;
		}
		const label = member.procKind === 'PropertyLet' ? 'Property Let' : 'Property Set';
		push(
			'propertySetterMissingValue',
			`${label} '${member.name}' must include a final value parameter.`,
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
			if (!reason) {
				continue;
			}
			push(
				'propertyAccessorSignatureMismatch',
				`${propertyProcedureLabel(setter.procKind)} '${setter.name}' argument list must match Property Get '${getter.name}' before the final value parameter. ${reason}`,
				declaredNameSpan(source, setter.span, setter.name),
			);
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
	}
	return undefined;
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

	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			inspect(decl.span, true);
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
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

function inspectTypeField(
	field: TypeFieldNode,
	inspect: (span: Span, allowEquals: boolean) => void,
): void {
	inspect(field.span, false);
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

export function checkInvalidAsTypeNames(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	opts: AnalyzeModuleOptions,
	push: PushFn,
): void {
	const withEventsNewDeclarationSpans = collectWithEventsNewDeclarationSpans(mod, activity);
	for (const ref of collectTypeNameReferences(source)) {
		if (activity?.isInactive(ref.span)) {
			continue;
		}
		const lookupName = typeReferenceLookupName(ref);
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
	}
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
				if (i !== params.length - 1) {
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
			const reason = parameterDefaultIncompatibilityReason(param, actual, memberCtx);
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
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		for (const param of member.params) {
			if (!param.defaultRaw) {
				continue;
			}
			if (resolveKnownObjectAssignmentType(param.asType, memberCtx)) {
				continue;
			}
			const defaultTokens = parameterDefaultTokens(source, param);
			if (!defaultTokens) {
				continue;
			}
			const nonConstant = nonConstantDefaultElement(defaultTokens.tokens, param.span.start);
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
 * are constant expressions and never flagged.
 */
export function checkNonConstantConstValues(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
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
			const nonConstant = nonConstantDefaultElement(valueTokens.tokens, decl.span.start);
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
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Enum') {
			continue;
		}
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
			const nonConstant = nonConstantDefaultElement(valueTokens.tokens, enumMember.span.start);
			if (!nonConstant) {
				continue;
			}
			push(
				'enumMemberNotConstant',
				`Enum member '${enumMember.name}' value must be a constant expression; ${nonConstant.label} is not constant.`,
				nonConstant.span,
			);
		}
	}
}

/**
 * Operator keywords (And, Or, Not, Mod, Xor, Eqv, Imp, Is, Like, TypeOf, plus
 * New/AddressOf) lex as `keyword` but are never callable names. They must be
 * excluded from the call heuristic below, otherwise a legal constant expression
 * like `6 And (3)` reads as a bogus call `And(...)`. New/AddressOf are still
 * flagged by the dedicated branch above this set's use.
 */
const OPERATOR_KEYWORD_WORDS = new Set(OPERATOR_IDENTIFIERS.map((word) => word.toLowerCase()));

function nonConstantDefaultElement(
	tokens: VbaToken[],
	baseOffset: number,
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
	memberCtx: MemberCompletionContext,
): string | undefined {
	if (param.isArray && isKnownScalarDefaultType(actual.type)) {
		return 'Optional array parameter defaults cannot be scalar values.';
	}
	const expectedRaw = param.asType;
	if (!expectedRaw) {
		return undefined;
	}
	const expectedObject = resolveKnownObjectAssignmentType(expectedRaw, memberCtx);
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
			return 'an Excel object-model type';
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
 * Rule: `Option` statements must precede every declaration and procedure (only
 * `Attribute` lines may come before them in an exported module). Once a real
 * declaration has appeared, any later `Option` is misplaced.
 */
export function checkOptionPlacement(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Declarations that precede the Option under test AND could be compiled
	// beside it: a declaration in the other arm of a chain closes no window,
	// because only one arm is ever built (issues/58).
	const declarationsAbove: Span[] = [];
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Attribute') {
			continue;
		}
		// A conditional-compilation directive is not a declaration, so it does
		// not close the window for Option statements. The live VBE compiles
		// `#Const FLAG = 1` above `Option Explicit` (oracle case
		// const_directive_before_option_explicit_compile), which the rule used
		// to report as a misplaced Option (issue #41).
		if (member.kind === 'ConditionalDirective') {
			continue;
		}
		if (member.kind === 'Option') {
			const compiledTogether = declarationsAbove.some(
				(prior) => !activity?.mutuallyExclusive(prior, member.span),
			);
			if (compiledTogether) {
				push(
					'optionAfterDeclaration',
					'Option statements must appear before any declaration or procedure.',
					firstTokenSpan(source, member.span),
				);
			}
			continue;
		}
		declarationsAbove.push(member.span);
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
	source: string,
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

/** VBA allows at most 60 parameters on a procedure. */
const MAX_PROCEDURE_PARAMETERS = 60;

/**
 * Rule: a procedure may declare at most 60 parameters. VBE rejects a 61st with
 * "Too many arguments" (oracle-verified `corpus_arg_limit_001b_compile`; 60 is
 * the documented VBA maximum).
 */
export function checkTooManyParameters(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.params.length <= MAX_PROCEDURE_PARAMETERS) {
			continue;
		}
		push(
			'tooManyParameters',
			`A procedure may have at most ${MAX_PROCEDURE_PARAMETERS} parameters; '${member.name}' declares ${member.params.length}.`,
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
	source: string,
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
 * reports only the Optional diagnostic (matching VBE).
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
		if (member.kind !== 'Procedure') {
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
