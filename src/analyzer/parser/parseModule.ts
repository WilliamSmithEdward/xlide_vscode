// Error-tolerant VBA parser producing the AST defined in nodes.ts.
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025):
//   - 4.2     Modules
//   - 5.2.1   Option <option-directive>
//   - 3.4     Conditional Compilation Directives (#Const / #If / #ElseIf / #Else / #End If)
//   - 5.2.3   Module Variable Declarations (Dim/Private/Public/Global/Static)
//   - 5.2.3.3 User Defined Types (Type ... End Type)
//   - 5.2.3.4 Enumerations (Enum ... End Enum)
//   - 5.2.3.5 External Procedure Declarations (Declare)
//   - 5.2.4   Const Declarations
//   - 5.3.1   Procedure Declarations (Sub / Function)
//   - 5.3.2   Property Declarations (Property Get / Let / Set)
//   - 5.3.1.x Parameter Lists (Optional / ByVal / ByRef / ParamArray / As)
//   - 5.4.2   Conditional statements (If ... End If)
//   - 5.4.2.4 Select Case ... End Select
//   - 5.4.2.5 For / For Each ... Next
//   - 5.4.2.x Do ... Loop, While ... Wend, With ... End With
//
// Design notes:
//   * The parser works over LOGICAL STATEMENTS (MS-VBAL 3.3.1 EOS), so its
//     natural recovery points are exactly the newline/colon boundaries the
//     roadmap mandates. It never throws on malformed input.
//   * Block statements track an "open block" stack so a stray terminator is
//     reported instead of corrupting the tree, and an unclosed block yields a
//     "missing End X" diagnostic while still returning a node (Phase 3
//     acceptance criteria).
//   * Bodies do not yet build a full expression AST; non-declaration statements
//     are captured as StatementNode with their raw text. Expression parsing is
//     a later phase.

import { VbaToken } from '../lexer/tokenKinds';
import { tokenize } from '../lexer/tokenize';
import {
	matchParenFrom,
	splitTopLevelTokenGroups,
	tokensWithoutLeadingLineNumber,
} from '../lexer/tokenHelpers';
import {
	AttributeNode,
	BodyNode,
	ConditionalDirectiveNode,
	DeclareNode,
	DoBlockNode,
	EventNode,
	EnumMemberNode,
	EnumNode,
	ForBlockNode,
	IfBlockNode,
	ModuleKind,
	ModuleMember,
	ModuleNode,
	OptionNode,
	ParameterNode,
	ParseDiagnostic,
	ParseSeverity,
	ProcedureNode,
	ProcKind,
	SelectBlockNode,
	Span,
	StatementNode,
	TypeFieldNode,
	TypeNode,
	VariableDeclNode,
	VariableGroupNode,
	WhileBlockNode,
	WithBlockNode,
} from './nodes';
import {
	codeTokens,
	LogicalStatement,
	splitLogicalStatements,
	StatementCursor,
	tokenWord,
} from './parserState';
import { parseFixedLengthStringType } from './fixedLengthString';
import {
	isTypeDeclarationSuffix,
	typeNameForDeclarationSuffix,
} from './typeDeclarationSuffix';

interface DeclaredNameInfo {
	name: string;
	nameSpan?: Span;
	nextIndex: number;
	typeSuffix?: string;
	typeSuffixSpan?: Span;
}

/** Visibility / sharing modifiers that may lead a declaration (MS-VBAL 5.2.3). */
const LEADING_MODIFIERS = new Set(['public', 'private', 'friend', 'global', 'static']);

/** Parameter mechanism markers (MS-VBAL 5.3.1.x). */
const PARAM_MARKERS = new Set(['optional', 'byval', 'byref', 'paramarray']);

/** Maps an expected-closer tag to its canonical label for diagnostics. */
const CLOSER_LABELS: Readonly<Record<string, string>> = {
	endif: 'End If',
	next: 'Next',
	loop: 'Loop',
	wend: 'Wend',
	endwith: 'End With',
	endselect: 'End Select',
	endsub: 'End Sub',
	endfunction: 'End Function',
	endproperty: 'End Property',
	endtype: 'End Type',
	endenum: 'End Enum',
};

// Editor surfaces (completion, hover, signature help, references) re-parse
// the same module text many times within one request, so a tiny value-keyed
// memo collapses those parses to one. The AST is treated as immutable by all
// consumers; callers must not mutate the returned nodes.
const PARSE_CACHE_MAX = 2;
const parseCache: { source: string; module: ModuleNode }[] = [];

/** Parse VBA source text into a ModuleNode AST. Never throws. */
export function parseModule(source: string): ModuleNode {
	for (let i = 0; i < parseCache.length; i += 1) {
		if (parseCache[i].source === source) {
			const hit = parseCache[i];
			if (i > 0) {
				parseCache.splice(i, 1);
				parseCache.unshift(hit);
			}
			return hit.module;
		}
	}
	const module = new Parser(source, tokenize(source)).parse();
	parseCache.unshift({ source, module });
	if (parseCache.length > PARSE_CACHE_MAX) {
		parseCache.pop();
	}
	return module;
}

class Parser {
	private readonly cursor: StatementCursor;
	private readonly diagnostics: ParseDiagnostic[] = [];
	/** Expected closers of the currently open blocks (innermost last). */
	private readonly openStack: string[] = [];

	constructor(
		private readonly source: string,
		tokens: readonly VbaToken[],
	) {
		this.cursor = new StatementCursor(splitLogicalStatements(tokens));
	}

	parse(): ModuleNode {
		const members: ModuleMember[] = [];
		while (!this.cursor.atEnd()) {
			const member = this.parseModuleMember();
			if (member) {
				members.push(member);
			}
		}
		const span: Span = {
			start: 0,
			end: this.source.length,
		};
		return {
			kind: 'Module',
			moduleKind: this.detectModuleKind(members),
			members,
			diagnostics: this.diagnostics,
			span,
		};
	}

	// -- Module level ------------------------------------------------------

	private parseModuleMember(): ModuleMember | undefined {
		const stmt = this.cursor.peek();
		if (!stmt) {
			return undefined;
		}
		const tokens = codeTokens(stmt);
		if (tokens.length === 0) {
			this.cursor.next();
			return undefined;
		}

		if (this.isAttribute(tokens)) {
			return this.parseAttribute(this.cursor.next()!, tokens);
		}
		if (this.isConditionalDirective(tokens)) {
			return this.parseConditionalDirective(this.cursor.next()!, tokens);
		}

		const modIndex = this.leadingModifierCount(tokens);
		const head = tokenWord(tokens[modIndex]);

		switch (head) {
			case 'option':
				return this.parseOption(this.cursor.next()!, tokens);
			case 'declare':
				return this.parseDeclare(this.cursor.next()!, tokens, modIndex);
			case 'event':
				return this.parseEvent(this.cursor.next()!, tokens, modIndex);
			case 'type':
				return this.parseTypeBlock(modIndex);
			case 'enum':
				return this.parseEnumBlock(modIndex);
			case 'sub':
			case 'function':
			case 'property':
				return this.parseProcedure();
			case 'const':
				return this.parseVariableGroup(this.cursor.next()!, tokens, modIndex, true);
			case 'dim':
				return this.parseVariableGroup(this.cursor.next()!, tokens, modIndex, false);
			default:
				// "Public Foo As Long": a module variable declared with only a
				// visibility modifier and no Dim keyword (MS-VBAL 5.2.3).
				if (modIndex > 0 && tokens[modIndex] !== undefined) {
					return this.parseVariableGroup(this.cursor.next()!, tokens, modIndex, false);
				}
				return this.makeStatement(this.cursor.next()!);
		}
	}

	private isAttribute(tokens: VbaToken[]): boolean {
		return tokens.length >= 1 && tokenWord(tokens[0]) === 'attribute';
	}

	private parseAttribute(stmt: LogicalStatement, tokens: VbaToken[]): AttributeNode {
		const eqIndex = tokens.findIndex((t) => t.rawText === '=');
		const nameStart = tokens[1];
		const nameEndIndex = eqIndex > 1 ? eqIndex - 1 : tokens.length - 1;
		const nameEnd = nameEndIndex >= 1 ? tokens[nameEndIndex] : undefined;
		const name =
			nameStart && nameEnd
				? tokens
					.slice(1, eqIndex > 1 ? eqIndex : tokens.length)
					.map((t) => t.rawText)
					.join('')
				: '';
		const valueRaw =
			eqIndex >= 0 && eqIndex + 1 < tokens.length
				? this.source.slice(tokens[eqIndex + 1].start, tokens[tokens.length - 1].end)
				: '';
		return {
			kind: 'Attribute',
			name,
			nameSpan: nameStart && nameEnd
				? { start: nameStart.start, end: nameEnd.end }
				: { start: stmt.start, end: stmt.start },
			valueRaw,
			span: { start: stmt.start, end: stmt.end },
		};
	}

	private parseOption(stmt: LogicalStatement, tokens: VbaToken[]): OptionNode {
		const optionText = tokens
			.slice(1)
			.map((t) => t.canonicalText ?? t.rawText)
			.join(' ');
		return {
			kind: 'Option',
			optionText,
			span: { start: stmt.start, end: stmt.end },
		};
	}

	private parseDeclare(
		stmt: LogicalStatement,
		tokens: VbaToken[],
		modIndex: number,
	): DeclareNode {
		const visibility = modIndex > 0 ? this.canonical(tokens[0]) : undefined;
		// Declare [PtrSafe] (Sub | Function) name ...
		const kindIndex = tokens.findIndex(
			(t) => tokenWord(t) === 'sub' || tokenWord(t) === 'function',
		);
		const isFunction = kindIndex >= 0 && tokenWord(tokens[kindIndex]) === 'function';
		const nameToken = kindIndex >= 0 ? tokens[kindIndex + 1] : undefined;
		const nameIndex = nameToken ? kindIndex + 1 : -1;
		const declaredName: DeclaredNameInfo = isFunction && nameIndex >= 0
			? this.parseDeclaredName(tokens, nameIndex, true)
			: {
				name: nameToken ? this.stripBrackets(nameToken.rawText) : '',
				nameSpan: this.tokenSpan(nameToken),
				nextIndex: nameIndex + 1,
			};
		const ptrSafe = kindIndex >= 0 && tokens
			.slice(modIndex + 1, kindIndex)
			.some((t) => tokenWord(t) === 'ptrsafe');
		const libIndex = nameIndex >= 0
			? tokens.findIndex((t, idx) => idx >= declaredName.nextIndex && tokenWord(t) === 'lib')
			: -1;
		const aliasIndex = nameIndex >= 0
			? tokens.findIndex((t, idx) => idx >= declaredName.nextIndex && tokenWord(t) === 'alias')
			: -1;
		const libName = libIndex >= 0 ? this.stringLiteralText(tokens[libIndex + 1]) : undefined;
		const aliasName = aliasIndex >= 0 ? this.stringLiteralText(tokens[aliasIndex + 1]) : undefined;

		let params: ParameterNode[] = [];
		let afterParen = declaredName.nextIndex;
		const lparen = nameIndex >= 0
			? tokens.findIndex((t, idx) => idx > nameIndex && t.rawText === '(')
			: -1;
		if (lparen >= 0) {
			const parsed = this.parseParamList(tokens, lparen);
			params = parsed.params;
			afterParen = parsed.closeIndex + 1;
		}

		let returnType: string | undefined;
		let hasAsClause = false;
		if (isFunction && tokens[afterParen] && tokenWord(tokens[afterParen]) === 'as') {
			hasAsClause = true;
			returnType = this.captureType(tokens, afterParen + 1);
		} else {
			returnType = typeNameForDeclarationSuffix(declaredName.typeSuffix);
		}
		return {
			kind: 'Declare',
			name: declaredName.name,
			nameSpan: declaredName.nameSpan,
			typeSuffix: declaredName.typeSuffix,
			typeSuffixSpan: declaredName.typeSuffixSpan,
			hasAsClause,
			isFunction,
			visibility,
			ptrSafe,
			libName,
			aliasName,
			params,
			returnType,
			span: { start: stmt.start, end: stmt.end },
		};
	}

	private parseEvent(
		stmt: LogicalStatement,
		tokens: VbaToken[],
		modIndex: number,
	): EventNode {
		const visibility = modIndex > 0 ? this.canonical(tokens[0]) : undefined;
		const nameToken = tokens[modIndex + 1];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		let params: ParameterNode[] = [];
		const lparen = nameToken
			? tokens.findIndex((t, idx) => idx > modIndex + 1 && t.rawText === '(')
			: -1;
		if (lparen >= 0) {
			params = this.parseParamList(tokens, lparen).params;
		}
		return {
			kind: 'Event',
			name,
			nameSpan: this.tokenSpan(nameToken),
			visibility,
			params,
			span: { start: stmt.start, end: stmt.end },
		};
	}

	private isConditionalDirective(tokens: VbaToken[]): boolean {
		return tokens[0]?.kind === 'directive';
	}

	private parseConditionalDirective(
		stmt: LogicalStatement,
		tokens: VbaToken[],
	): ConditionalDirectiveNode {
		const directiveWord = tokenWord(tokens[1]);
		const base = {
			kind: 'ConditionalDirective' as const,
			span: { start: stmt.start, end: stmt.end },
		};
		switch (directiveWord) {
			case 'const': {
				const nameToken = tokens[2];
				const eqIndex = tokens.findIndex((t, idx) => idx > 2 && t.rawText === '=');
				const value = eqIndex >= 0
					? this.tokenRangeRaw(tokens, eqIndex + 1, tokens.length)
					: {};
				return {
					...base,
					directiveKind: 'Const',
					name: nameToken ? this.stripBrackets(nameToken.rawText) : undefined,
					nameSpan: nameToken ? { start: nameToken.start, end: nameToken.end } : undefined,
					valueRaw: value.raw,
					valueSpan: value.span,
				};
			}
			case 'if': {
				const condition = this.directiveCondition(tokens, 2);
				return {
					...base,
					directiveKind: 'If',
					conditionRaw: condition.raw,
					conditionSpan: condition.span,
				};
			}
			case 'elseif': {
				const condition = this.directiveCondition(tokens, 2);
				return {
					...base,
					directiveKind: 'ElseIf',
					conditionRaw: condition.raw,
					conditionSpan: condition.span,
				};
			}
			case 'else':
				return {
					...base,
					directiveKind: 'Else',
				};
			case 'end':
				if (tokenWord(tokens[2]) === 'if') {
					return {
						...base,
						directiveKind: 'EndIf',
					};
				}
				break;
			case 'endif':
				return {
					...base,
					directiveKind: 'EndIf',
				};
		}
		const unknown = this.tokenRangeRaw(tokens, 1, tokens.length);
		return {
			...base,
			directiveKind: 'Unknown',
			conditionRaw: unknown.raw,
			conditionSpan: unknown.span,
		};
	}

	private parseVariableGroup(
		stmt: LogicalStatement,
		tokens: VbaToken[],
		modIndex: number,
		isConst: boolean,
	): VariableGroupNode {
		// Skip modifiers and the Dim/Const keyword (if present) to reach the names.
		let i = modIndex;
		const head = tokenWord(tokens[i]);
		const modifier = modIndex > 0 ? this.canonical(tokens[0]) : head === 'dim' ? 'Dim' : '';
		if (head === 'dim' || head === 'const') {
			i++;
		}
		let withEvents = false;
		if (tokenWord(tokens[i]) === 'withevents') {
			withEvents = true;
			i++;
		}
		const declarations = this.parseDeclaratorList(tokens, i, isConst);
		return {
			kind: 'VariableGroup',
			modifier,
			isConst,
			withEvents,
			declarations,
			span: { start: stmt.start, end: stmt.end },
		};
	}

	/** Parse a comma-separated declarator list: name[()][As type][= value], ... */
	private parseDeclaratorList(
		tokens: VbaToken[],
		from: number,
		isConst: boolean,
	): VariableDeclNode[] {
		const groups = this.splitTopLevelCommas(tokens, from, tokens.length);
		const declarations: VariableDeclNode[] = [];
		for (const group of groups) {
			if (group.length === 0) {
				continue;
			}
			declarations.push(this.parseDeclarator(group, isConst));
		}
		return declarations;
	}

	private parseDeclarator(group: VbaToken[], isConst: boolean): VariableDeclNode {
		let i = 0;
		const declaredName = this.parseDeclaredName(group, i, true);
		const name = declaredName.name;
		i = declaredName.nextIndex;
		let isArray = false;
		let arrayBounds: string | undefined;
		if (group[i] && group[i].rawText === '(') {
			isArray = true;
			const close = this.matchParen(group, i);
			if (group[close]?.rawText === ')') {
				const bounds = this.source.slice(group[i].end, group[close].start).trim();
				arrayBounds = bounds || undefined;
			}
			i = close + 1;
		}
		let isNew = false;
		let asType: string | undefined;
		let fixedLength: string | undefined;
		let hasAsClause = false;
		if (group[i] && tokenWord(group[i]) === 'as') {
			hasAsClause = true;
			i++;
			if (group[i] && tokenWord(group[i]) === 'new') {
				isNew = true;
				i++;
			}
			const type = this.captureDeclarationType(group, i, isConst ? '=' : undefined);
			asType = type.asType;
			fixedLength = type.fixedLength;
		} else {
			asType = typeNameForDeclarationSuffix(declaredName.typeSuffix);
		}
		let defaultRaw: string | undefined;
		const eq = group.findIndex((t) => t.rawText === '=');
		if (eq >= 0 && eq + 1 < group.length) {
			defaultRaw = this.source.slice(group[eq + 1].start, group[group.length - 1].end);
		}
		const first = group[0];
		const last = group[group.length - 1];
		return {
			kind: 'VariableDecl',
			name,
			nameSpan: declaredName.nameSpan,
			typeSuffix: declaredName.typeSuffix,
			typeSuffixSpan: declaredName.typeSuffixSpan,
			hasAsClause,
			asType,
			fixedLength,
			defaultRaw,
			isArray,
			arrayBounds,
			isNew,
			span: { start: first.start, end: last.end },
		};
	}

	// -- Type / Enum -------------------------------------------------------

	private parseTypeBlock(modIndex: number): TypeNode {
		const head = this.cursor.next()!;
		const tokens = codeTokens(head);
		const visibility = modIndex > 0 ? this.canonical(tokens[0]) : undefined;
		const nameToken = tokens[modIndex + 1];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		const fields: TypeFieldNode[] = [];
		let closed = false;
		let endStmt: LogicalStatement | undefined;
		while (!this.cursor.atEnd()) {
			const stmt = this.cursor.peek()!;
			if (this.closerKind(stmt) === 'endtype') {
				endStmt = this.cursor.next();
				closed = true;
				break;
			}
			if (this.isModuleLevelStarter(stmt) && !this.isTypeFieldStatement(stmt)) {
				break;
			}
			this.cursor.next();
			const ftokens = codeTokens(stmt);
			if (ftokens.length > 0) {
				fields.push(this.parseTypeField(stmt, ftokens));
			}
		}
		if (!closed) {
			this.diag(head, 'Type block is missing End Type.', 'error', 'MS-VBAL 5.2.3.3');
		}
		return {
			kind: 'Type',
			name,
			nameSpan: this.tokenSpan(nameToken),
			visibility,
			fields,
			closed,
			span: { start: head.start, end: (endStmt ?? head).end },
		};
	}

	private parseTypeField(stmt: LogicalStatement, tokens: VbaToken[]): TypeFieldNode {
		const declaredName = this.parseDeclaredName(tokens, 0, true);
		const name = declaredName.name;
		let i = declaredName.nextIndex;
		let isArray = false;
		if (tokens[i] && tokens[i].rawText === '(') {
			isArray = true;
			i = this.skipParens(tokens, i);
		}
		let asType: string | undefined;
		let fixedLength: string | undefined;
		let hasAsClause = false;
		if (tokens[i] && tokenWord(tokens[i]) === 'as') {
			hasAsClause = true;
			const type = this.captureDeclarationType(tokens, i + 1);
			asType = type.asType;
			fixedLength = type.fixedLength;
		} else {
			asType = typeNameForDeclarationSuffix(declaredName.typeSuffix);
		}
		return {
			kind: 'TypeField',
			name,
			nameSpan: declaredName.nameSpan,
			typeSuffix: declaredName.typeSuffix,
			typeSuffixSpan: declaredName.typeSuffixSpan,
			hasAsClause,
			asType,
			fixedLength,
			isArray,
			span: { start: stmt.start, end: stmt.end },
		};
	}

	private isTypeFieldStatement(stmt: LogicalStatement): boolean {
		const tokens = codeTokens(stmt);
		if (tokens.length === 0) {
			return false;
		}
		const first = tokenWord(tokens[0]);
		if (first === 'type' && tokenWord(tokens[1]) === 'as') {
			return true;
		}
		return tokens.some((token, index) => index > 0 && tokenWord(token) === 'as');
	}

	private parseEnumBlock(modIndex: number): EnumNode {
		const head = this.cursor.next()!;
		const tokens = codeTokens(head);
		const visibility = modIndex > 0 ? this.canonical(tokens[0]) : undefined;
		const nameToken = tokens[modIndex + 1];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		const members: EnumMemberNode[] = [];
		let closed = false;
		let endStmt: LogicalStatement | undefined;
		while (!this.cursor.atEnd()) {
			const stmt = this.cursor.peek()!;
			if (this.closerKind(stmt) === 'endenum') {
				endStmt = this.cursor.next();
				closed = true;
				break;
			}
			if (this.isModuleLevelStarter(stmt)) {
				break;
			}
			this.cursor.next();
			const mtokens = codeTokens(stmt);
			if (mtokens.length > 0) {
				const eqIndex = mtokens.findIndex((t) => t.rawText === '=');
				const valueRaw =
					eqIndex >= 0 && eqIndex + 1 < mtokens.length
						? this.source.slice(mtokens[eqIndex + 1].start, mtokens[mtokens.length - 1].end)
						: undefined;
				members.push({
					kind: 'EnumMember',
					name: this.stripBrackets(mtokens[0].rawText),
					nameSpan: this.tokenSpan(mtokens[0]),
					valueRaw,
					span: { start: stmt.start, end: stmt.end },
				});
			}
		}
		if (!closed) {
			this.diag(head, 'Enum block is missing End Enum.', 'error', 'MS-VBAL 5.2.3.4');
		}
		return {
			kind: 'Enum',
			name,
			nameSpan: this.tokenSpan(nameToken),
			visibility,
			members,
			closed,
			span: { start: head.start, end: (endStmt ?? head).end },
		};
	}

	// -- Procedures --------------------------------------------------------

	private parseProcedure(): ProcedureNode {
		const head = this.cursor.next()!;
		const tokens = codeTokens(head);
		const modIndex = this.leadingModifierCount(tokens);
		const modifiers = tokens.slice(0, modIndex).map((t) => this.canonical(t));

		let i = modIndex;
		const headWord = tokenWord(tokens[i]);
		let procKind: ProcKind;
		if (headWord === 'property') {
			i++;
			const accessor = tokenWord(tokens[i]);
			procKind =
				accessor === 'get'
					? 'PropertyGet'
					: accessor === 'set'
						? 'PropertySet'
						: 'PropertyLet';
			i++;
		} else if (headWord === 'function') {
			procKind = 'Function';
			i++;
		} else {
			procKind = 'Sub';
			i++;
		}

		const declaredName = this.parseDeclaredName(
			tokens,
			i,
			procKind === 'Function' || procKind === 'PropertyGet',
		);
		const name = declaredName.name;
		i = declaredName.nextIndex;

		let params: ParameterNode[] = [];
		let afterParen = i;
		if (tokens[i] && tokens[i].rawText === '(') {
			const parsed = this.parseParamList(tokens, i);
			params = parsed.params;
			afterParen = parsed.closeIndex + 1;
		}

		let returnType: string | undefined;
		let hasAsClause = false;
		if (tokens[afterParen] && tokenWord(tokens[afterParen]) === 'as') {
			hasAsClause = true;
			returnType = this.captureType(tokens, afterParen + 1);
		} else {
			returnType = typeNameForDeclarationSuffix(declaredName.typeSuffix);
		}

		const expected = this.procCloser(procKind);
		this.openStack.push(expected);
		const body: BodyNode[] = [];
		const attributes: AttributeNode[] = [];
		let closed = false;
		let endStmt: LogicalStatement | undefined;
		let sawConditionalDirective = false;
		while (!this.cursor.atEnd()) {
			const stmt = this.cursor.peek()!;
			const ck = this.closerKind(stmt);
			if (ck === expected) {
				endStmt = this.cursor.next();
				closed = true;
				break;
			}
			const stmtTokens = codeTokens(stmt);
			if (this.isAttribute(stmtTokens)) {
				if (this.isExportedProcedureAttribute(stmt, stmtTokens, name, body.length === 0)) {
					attributes.push(this.parseAttribute(this.cursor.next()!, stmtTokens));
					continue;
				}
				const item = this.parseBodyItem(stmt);
				if (item === undefined) {
					break;
				}
				body.push(item);
				continue;
			}
			const nestedModuleBlock = this.nestedTypeOrEnumBlockKind(stmt);
			if (nestedModuleBlock) {
				body.push(this.parseInvalidNestedModuleBlockStatement(nestedModuleBlock, [expected]));
				continue;
			}
			// Recovery: a new module-level construct means the End was forgotten.
			if (this.isModuleLevelStarter(stmt)) {
				if (
					sawConditionalDirective &&
					this.isAlternativeProcedureHeader(stmt, procKind, name)
				) {
					this.cursor.next();
					continue;
				}
				break;
			}
			const item = this.parseBodyItem(stmt);
			if (item === undefined) {
				break;
			}
			if (item.kind === 'ConditionalDirective') {
				sawConditionalDirective = true;
			}
			body.push(item);
		}
		this.openStack.pop();
		if (!closed) {
			this.diag(
				head,
				`Procedure '${name}' is missing ${CLOSER_LABELS[expected]}.`,
				'error',
				'MS-VBAL 5.3.1',
			);
		}
		const lastBody = body[body.length - 1];
		const end = endStmt ? endStmt.end : lastBody ? lastBody.span.end : head.end;
		return {
			kind: 'Procedure',
			procKind,
			name,
			nameSpan: declaredName.nameSpan,
			typeSuffix: declaredName.typeSuffix,
			typeSuffixSpan: declaredName.typeSuffixSpan,
			hasAsClause,
			modifiers,
			params,
			returnType,
			...(attributes.length > 0 ? { attributes } : {}),
			body,
			closed,
			span: { start: head.start, end },
		};
	}

	private procCloser(kind: ProcKind): string {
		if (kind === 'Function') {
			return 'endfunction';
		}
		if (kind === 'Sub') {
			return 'endsub';
		}
		return 'endproperty';
	}

	private isAlternativeProcedureHeader(
		stmt: LogicalStatement,
		currentKind: ProcKind,
		currentName: string,
	): boolean {
		const tokens = codeTokens(stmt);
		const modIndex = this.leadingModifierCount(tokens);
		let i = modIndex;
		const headWord = tokenWord(tokens[i]);
		let kind: ProcKind | undefined;
		if (headWord === 'property') {
			i++;
			const accessor = tokenWord(tokens[i]);
			kind =
				accessor === 'get'
					? 'PropertyGet'
					: accessor === 'set'
						? 'PropertySet'
						: accessor === 'let'
							? 'PropertyLet'
							: undefined;
			i++;
		} else if (headWord === 'function') {
			kind = 'Function';
			i++;
		} else if (headWord === 'sub') {
			kind = 'Sub';
			i++;
		}
		const nameToken = tokens[i];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		return kind === currentKind && name.toLowerCase() === currentName.toLowerCase();
	}

	private parseParamList(
		tokens: VbaToken[],
		lparen: number,
	): { params: ParameterNode[]; closeIndex: number } {
		const closeIndex = this.matchParen(tokens, lparen);
		const groups = this.splitTopLevelCommas(tokens, lparen + 1, closeIndex);
		const params = groups
			.filter((g) => g.length > 0)
			.map((g) => this.parseParam(g));
		return { params, closeIndex };
	}

	private parseParam(group: VbaToken[]): ParameterNode {
		let i = 0;
		let optional = false;
		let byVal = false;
		let byRef = false;
		let paramArray = false;
		while (group[i] && PARAM_MARKERS.has(tokenWord(group[i]))) {
			switch (tokenWord(group[i])) {
				case 'optional':
					optional = true;
					break;
				case 'byval':
					byVal = true;
					break;
				case 'byref':
					byRef = true;
					break;
				case 'paramarray':
					paramArray = true;
					break;
			}
			i++;
		}
		const declaredName = this.parseDeclaredName(group, i, true);
		const name = declaredName.name;
		i = declaredName.nextIndex;
		let isArray = false;
		if (group[i] && group[i].rawText === '(') {
			isArray = true;
			i = this.skipParens(group, i);
		}
		let asType: string | undefined;
		let hasAsClause = false;
		if (group[i] && tokenWord(group[i]) === 'as') {
			hasAsClause = true;
			i++;
			asType = this.captureType(group, i, '=');
		} else {
			asType = typeNameForDeclarationSuffix(declaredName.typeSuffix);
		}
		let defaultRaw: string | undefined;
		const eq = group.findIndex((t) => t.rawText === '=');
		if (eq >= 0 && eq + 1 < group.length) {
			defaultRaw = this.source.slice(group[eq + 1].start, group[group.length - 1].end);
		}
		const first = group[0];
		const last = group[group.length - 1];
		return {
			kind: 'Parameter',
			name,
			nameSpan: declaredName.nameSpan,
			typeSuffix: declaredName.typeSuffix,
			typeSuffixSpan: declaredName.typeSuffixSpan,
			hasAsClause,
			optional,
			byVal,
			byRef,
			paramArray,
			asType,
			isArray,
			defaultRaw,
			span: { start: first.start, end: last.end },
		};
	}

	// -- Procedure body items / block statements ---------------------------

	/** Parse one statement inside a procedure/block body. */
	private parseBodyItem(stmt: LogicalStatement): BodyNode | undefined {
		const ck = this.closerKind(stmt);
		if (ck) {
			if (this.openStack.includes(ck)) {
				// Belongs to an ancestor block; stop and let it close.
				return undefined;
			}
			// A terminator with no matching open block.
			this.diag(
				stmt,
				`Unexpected '${CLOSER_LABELS[ck]}' without a matching opening block.`,
				'error',
				'MS-VBAL 5.4',
			);
			this.cursor.next();
			return this.makeStatement(stmt);
		}
		const opener = this.openerKind(stmt);
		if (opener) {
			return this.parseBlock(opener);
		}
		const tokens = codeTokensAfterLineNumber(stmt);
		if (this.isConditionalDirective(tokens)) {
			this.cursor.next();
			return this.parseConditionalDirective(stmt, tokens);
		}
		const head = tokenWord(tokens[0]);
		if (head === 'dim' || head === 'const' || head === 'static') {
			this.cursor.next();
			const modIndex = head === 'static' ? 1 : 0;
			return this.parseVariableGroup(stmt, tokens, modIndex, head === 'const');
		}
		this.cursor.next();
		return this.makeStatement(stmt);
	}

	private parseInvalidNestedModuleBlockStatement(
		kind: 'type' | 'enum',
		stopClosers: readonly string[],
	): StatementNode {
		const head = this.cursor.next()!;
		const expected = kind === 'type' ? 'endtype' : 'endenum';
		let end = head.end;
		while (!this.cursor.atEnd()) {
			const stmt = this.cursor.peek()!;
			const closer = this.closerKind(stmt);
			if (closer === expected) {
				end = this.cursor.next()!.end;
				break;
			}
			if (closer && stopClosers.includes(closer)) {
				break;
			}
			if (
				this.isModuleLevelStarter(stmt) &&
				!(kind === 'type' && this.isTypeFieldStatement(stmt))
			) {
				break;
			}
			end = this.cursor.next()!.end;
		}
		return {
			kind: 'Statement',
			raw: this.source.slice(head.start, end),
			span: { start: head.start, end },
		};
	}

	private parseBlock(
		opener: 'if' | 'for' | 'foreach' | 'do' | 'while' | 'with' | 'select',
	): BodyNode {
		const head = this.cursor.next()!;
		const expected = this.blockCloser(opener);
		this.openStack.push(expected);
		const body: BodyNode[] = [];
		let closed = false;
		let endStmt: LogicalStatement | undefined;
		while (!this.cursor.atEnd()) {
			const stmt = this.cursor.peek()!;
			const ck = this.closerKind(stmt);
			if (ck === expected) {
				endStmt = this.cursor.next();
				closed = true;
				break;
			}
			const nestedModuleBlock = this.nestedTypeOrEnumBlockKind(stmt);
			if (nestedModuleBlock) {
				body.push(this.parseInvalidNestedModuleBlockStatement(nestedModuleBlock, [expected]));
				continue;
			}
			if (this.isModuleLevelStarter(stmt)) {
				break;
			}
			const item = this.parseBodyItem(stmt);
			if (item === undefined) {
				break;
			}
			body.push(item);
		}
		this.openStack.pop();
		if (!closed) {
			this.diag(
				head,
				`Block is missing ${CLOSER_LABELS[expected]}.`,
				'error',
				'MS-VBAL 5.4',
			);
		}
		const span: Span = { start: head.start, end: (endStmt ?? head).end };
		return this.makeBlockNode(opener, body, closed, span, head, endStmt);
	}

	private makeBlockNode(
		opener: 'if' | 'for' | 'foreach' | 'do' | 'while' | 'with' | 'select',
		body: BodyNode[],
		closed: boolean,
		span: Span,
		head: LogicalStatement,
		endStmt: LogicalStatement | undefined,
	): BodyNode {
		switch (opener) {
			case 'if':
				return { kind: 'IfBlock', body, closed, span } satisfies IfBlockNode;
			case 'for':
			case 'foreach': {
				const control = this.forControlVariable(opener, head);
				const source = opener === 'foreach' ? this.forEachSourceExpression(head) : undefined;
				const next = this.nextControlVariable(endStmt);
				return {
					kind: 'ForBlock',
					each: opener === 'foreach',
					...(control ? {
						controlVariable: control.name,
						controlVariableSpan: control.span,
					} : {}),
					...(source ? {
						sourceExpression: source.raw,
						sourceExpressionSpan: source.span,
					} : {}),
					...(next ? {
						nextVariable: next.name,
						nextVariableSpan: next.span,
					} : {}),
					body,
					closed,
					span,
				} satisfies ForBlockNode;
			}
			case 'do':
				return { kind: 'DoBlock', body, closed, span } satisfies DoBlockNode;
			case 'while':
				return { kind: 'WhileBlock', body, closed, span } satisfies WhileBlockNode;
			case 'with':
				return { kind: 'WithBlock', body, closed, span } satisfies WithBlockNode;
			case 'select':
				return { kind: 'SelectBlock', body, closed, span } satisfies SelectBlockNode;
		}
	}

	private forControlVariable(
		opener: 'for' | 'foreach',
		stmt: LogicalStatement,
	): { name: string; span: Span } | undefined {
		const tokens = codeTokensAfterLineNumber(stmt);
		const index = opener === 'foreach' ? 2 : 1;
		const nameToken = tokens[index];
		const name = this.simpleNameFromToken(nameToken);
		if (!name) {
			return undefined;
		}
		if (opener === 'foreach') {
			if (tokenWord(tokens[index + 1]) !== 'in') {
				return undefined;
			}
		} else if (tokens[index + 1]?.rawText !== '=') {
			return undefined;
		}
		return { name, span: { start: nameToken.start, end: nameToken.end } };
	}

	private forEachSourceExpression(
		stmt: LogicalStatement,
	): { raw: string; span: Span } | undefined {
		const tokens = codeTokensAfterLineNumber(stmt);
		const inIndex = tokenWord(tokens[3]) === 'in' ? 3 : -1;
		if (inIndex < 0 || inIndex + 1 >= tokens.length) {
			return undefined;
		}
		const first = tokens[inIndex + 1];
		const last = tokens[tokens.length - 1];
		return {
			raw: this.source.slice(first.start, last.end),
			span: { start: first.start, end: last.end },
		};
	}

	private nextControlVariable(
		stmt: LogicalStatement | undefined,
	): { name: string; span: Span } | undefined {
		if (!stmt) {
			return undefined;
		}
		const tokens = codeTokensAfterLineNumber(stmt);
		if (tokenWord(tokens[0]) !== 'next' || tokens.length !== 2) {
			return undefined;
		}
		const nameToken = tokens[1];
		const name = this.simpleNameFromToken(nameToken);
		return name ? { name, span: { start: nameToken.start, end: nameToken.end } } : undefined;
	}

	private simpleNameFromToken(token: VbaToken | undefined): string | undefined {
		if (!token) {
			return undefined;
		}
		if (token.kind === 'identifier' || token.kind === 'keyword') {
			return token.rawText;
		}
		if (token.kind === 'bracketedIdentifier') {
			return this.stripBrackets(token.rawText);
		}
		return undefined;
	}

	private blockCloser(
		opener: 'if' | 'for' | 'foreach' | 'do' | 'while' | 'with' | 'select',
	): string {
		switch (opener) {
			case 'if':
				return 'endif';
			case 'for':
			case 'foreach':
				return 'next';
			case 'do':
				return 'loop';
			case 'while':
				return 'wend';
			case 'with':
				return 'endwith';
			case 'select':
				return 'endselect';
		}
	}

	/** Detect a block-opening statement (MS-VBAL 5.4). */
	private openerKind(
		stmt: LogicalStatement,
	): 'if' | 'for' | 'foreach' | 'do' | 'while' | 'with' | 'select' | undefined {
		const tokens = codeTokensAfterLineNumber(stmt);
		const w0 = tokenWord(tokens[0]);
		switch (w0) {
			case 'if': {
				// Multi-line If only when "Then" is the final code token; a
				// single-line "If x Then stmt" is not a block (MS-VBAL 5.4.2.1).
				const last = tokens[tokens.length - 1];
				return tokenWord(last) === 'then' ? 'if' : undefined;
			}
			case 'for':
				return tokenWord(tokens[1]) === 'each' ? 'foreach' : 'for';
			case 'do':
				return 'do';
			case 'while':
				return 'while';
			case 'with':
				return 'with';
			case 'select':
				return tokenWord(tokens[1]) === 'case' ? 'select' : undefined;
			default:
				return undefined;
		}
	}

	/** Detect a block-closing statement and return its expected-closer tag. */
	private closerKind(stmt: LogicalStatement): string | undefined {
		const tokens = codeTokensAfterLineNumber(stmt);
		const w0 = tokenWord(tokens[0]);
		if (w0 === 'next') {
			return 'next';
		}
		if (w0 === 'loop') {
			return 'loop';
		}
		if (w0 === 'wend') {
			return 'wend';
		}
		if (w0 === 'end') {
			// "End" alone (MS-VBAL 5.4.7) is a statement, not a block closer.
			const w1 = tokenWord(tokens[1]);
			switch (w1) {
				case 'if':
					return 'endif';
				case 'with':
					return 'endwith';
				case 'select':
					return 'endselect';
				case 'sub':
					return 'endsub';
				case 'function':
					return 'endfunction';
				case 'property':
					return 'endproperty';
				case 'type':
					return 'endtype';
				case 'enum':
					return 'endenum';
				default:
					return undefined;
			}
		}
		return undefined;
	}

	/** A statement that should only appear at module level (recovery boundary). */
	private nestedTypeOrEnumBlockKind(stmt: LogicalStatement): 'type' | 'enum' | undefined {
		const tokens = codeTokensAfterLineNumber(stmt);
		const modIndex = this.leadingModifierCount(tokens);
		const head = tokenWord(tokens[modIndex]);
		return head === 'type' || head === 'enum' ? head : undefined;
	}

	private isModuleLevelStarter(stmt: LogicalStatement): boolean {
		const tokens = codeTokens(stmt);
		const modIndex = this.leadingModifierCount(tokens);
		const head = tokenWord(tokens[modIndex]);
		switch (head) {
			case 'sub':
			case 'function':
			case 'property':
			case 'type':
			case 'enum':
			case 'declare':
				return true;
			default:
				return tokenWord(tokens[0]) === 'attribute';
		}
	}

	private isExportedProcedureAttribute(
		stmt: LogicalStatement,
		tokens: VbaToken[],
		procedureName: string,
		inMemberMetadataSlot: boolean,
	): boolean {
		if (!this.startsAtPhysicalLineStart(stmt)) {
			return false;
		}
		if (!inMemberMetadataSlot) {
			return false;
		}
		const eqIndex = tokens.findIndex((t) => t.rawText === '=');
		if (eqIndex <= 1) {
			return false;
		}
		const attrName = tokens
			.slice(1, eqIndex)
			.map((token) => token.rawText)
			.join('');
		const dot = attrName.indexOf('.');
		if (dot <= 0) {
			return false;
		}
		const target = this.stripBrackets(attrName.slice(0, dot));
		const memberAttributeName = attrName.slice(dot + 1);
		return (
			target.toLowerCase() === procedureName.toLowerCase() &&
			/^VB_[A-Za-z0-9_]+$/i.test(memberAttributeName)
		);
	}

	private startsAtPhysicalLineStart(stmt: LogicalStatement): boolean {
		const previousNewline = Math.max(
			this.source.lastIndexOf('\n', stmt.start - 1),
			this.source.lastIndexOf('\r', stmt.start - 1),
		);
		return stmt.start === previousNewline + 1;
	}

	// -- Token helpers -----------------------------------------------------

	private leadingModifierCount(tokens: VbaToken[]): number {
		let i = 0;
		while (tokens[i] && LEADING_MODIFIERS.has(tokenWord(tokens[i]))) {
			i++;
		}
		return i;
	}

	/** Split tokens[from, to) into comma-separated groups at paren depth 0. */
	private splitTopLevelCommas(tokens: VbaToken[], from: number, to: number): VbaToken[][] {
		return splitTopLevelTokenGroups(tokens, from, ',', to);
	}

	private directiveCondition(
		tokens: VbaToken[],
		from: number,
	): { raw?: string; span?: Span } {
		const to = tokenWord(tokens[tokens.length - 1]) === 'then'
			? tokens.length - 1
			: tokens.length;
		return this.tokenRangeRaw(tokens, from, to);
	}

	private tokenRangeRaw(
		tokens: VbaToken[],
		from: number,
		to: number,
	): { raw?: string; span?: Span } {
		if (from < 0 || from >= to || from >= tokens.length) {
			return {};
		}
		const startToken = tokens[from];
		const endToken = tokens[Math.min(to, tokens.length) - 1];
		if (!startToken || !endToken) {
			return {};
		}
		return {
			raw: this.source.slice(startToken.start, endToken.end),
			span: { start: startToken.start, end: endToken.end },
		};
	}

	/** Index of the matching ')', or the last token when unmatched. */
	private matchParen(tokens: VbaToken[], lparen: number): number {
		const close = matchParenFrom(tokens, lparen);
		return close >= 0 ? close : tokens.length - 1;
	}

	/** Index just past a parenthesized group starting at index i (on '('). */
	private skipParens(tokens: VbaToken[], i: number): number {
		return this.matchParen(tokens, i) + 1;
	}

	/** Capture a type expression starting at index i up to a terminator. */
	private captureType(tokens: VbaToken[], i: number, stopRaw?: string): string | undefined {
		if (!tokens[i]) {
			return undefined;
		}
		let depth = 0;
		let last = i;
		for (let j = i; j < tokens.length; j++) {
			const t = tokens[j];
			if (t.rawText === '(') {
				depth++;
			} else if (t.rawText === ')') {
				if (depth === 0) {
					break;
				}
				depth--;
			}
			if (depth === 0 && (t.rawText === ',' || (stopRaw && t.rawText === stopRaw))) {
				break;
			}
			last = j;
		}
		return this.source.slice(tokens[i].start, tokens[last].end);
	}

	private captureDeclarationType(
		tokens: VbaToken[],
		i: number,
		stopRaw?: string,
	): { asType?: string; fixedLength?: string } {
		const fixed = parseFixedLengthStringType(tokens, i);
		if (fixed) {
			return {
				asType: this.source.slice(tokens[i].start, tokens[i].end),
				fixedLength: this.source.slice(
					tokens[fixed.lengthIndex].start,
					tokens[fixed.lengthIndex].end,
				),
			};
		}
		return { asType: this.captureType(tokens, i, stopRaw) };
	}

	private canonical(token: VbaToken | undefined): string {
		if (!token) {
			return '';
		}
		return token.canonicalText ?? token.rawText;
	}

	private parseDeclaredName(
		tokens: VbaToken[],
		index: number,
		allowTypeSuffix: boolean,
	): DeclaredNameInfo {
		const nameToken = tokens[index];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		const nameSpan = this.tokenSpan(nameToken);
		const suffixToken = allowTypeSuffix ? tokens[index + 1] : undefined;
		if (
			nameToken &&
			suffixToken &&
			nameToken.end === suffixToken.start &&
			isTypeDeclarationSuffix(suffixToken.rawText)
		) {
			return {
				name,
				nameSpan,
				nextIndex: index + 2,
				typeSuffix: suffixToken.rawText,
				typeSuffixSpan: { start: suffixToken.start, end: suffixToken.end },
			};
		}
		return { name, nameSpan, nextIndex: index + 1 };
	}

	/** Span of one token, or undefined when the token is missing. */
	private tokenSpan(token: VbaToken | undefined): Span | undefined {
		return token ? { start: token.start, end: token.end } : undefined;
	}

	private stripBrackets(raw: string): string {
		if (raw.length >= 2 && raw.startsWith('[') && raw.endsWith(']')) {
			return raw.slice(1, -1);
		}
		return raw;
	}

	private stringLiteralText(token: VbaToken | undefined): string | undefined {
		if (!token || token.kind !== 'stringLiteral') {
			return undefined;
		}
		const raw = token.rawText;
		if (raw.length >= 2 && raw.startsWith('"') && raw.endsWith('"')) {
			return raw.slice(1, -1).replace(/""/g, '"');
		}
		return raw;
	}

	private makeStatement(stmt: LogicalStatement): StatementNode {
		return {
			kind: 'Statement',
			raw: this.source.slice(stmt.start, stmt.end),
			span: { start: stmt.start, end: stmt.end },
		};
	}

	private detectModuleKind(members: ModuleMember[]): ModuleKind {
		for (const m of members) {
			if (m.kind === 'Attribute' && /^VB_(Exposed|Creatable|PredeclaredId)$/i.test(m.name)) {
				return 'class';
			}
		}
		return 'unknown';
	}

	private diag(
		at: LogicalStatement,
		message: string,
		severity: ParseSeverity,
		specRef?: string,
	): void {
		this.diagnostics.push({
			span: { start: at.start, end: at.end },
			message,
			severity,
			specRef,
		});
	}
}

function codeTokensAfterLineNumber(statement: LogicalStatement): VbaToken[] {
	return tokensWithoutLeadingLineNumber(codeTokens(statement));
}
