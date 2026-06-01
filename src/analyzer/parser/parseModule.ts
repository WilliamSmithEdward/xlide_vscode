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
	AttributeNode,
	BodyNode,
	ConditionalDirectiveNode,
	DeclareNode,
	DoBlockNode,
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

/** Parse VBA source text into a ModuleNode AST. Never throws. */
export function parseModule(source: string): ModuleNode {
	const tokens = tokenize(source);
	return new Parser(source, tokens).parse();
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
		const ptrSafe = kindIndex >= 0 && tokens
			.slice(modIndex + 1, kindIndex)
			.some((t) => tokenWord(t) === 'ptrsafe');
		const libIndex = nameIndex >= 0
			? tokens.findIndex((t, idx) => idx > nameIndex && tokenWord(t) === 'lib')
			: -1;
		const aliasIndex = nameIndex >= 0
			? tokens.findIndex((t, idx) => idx > nameIndex && tokenWord(t) === 'alias')
			: -1;
		const libName = libIndex >= 0 ? this.stringLiteralText(tokens[libIndex + 1]) : undefined;
		const aliasName = aliasIndex >= 0 ? this.stringLiteralText(tokens[aliasIndex + 1]) : undefined;

		let params: ParameterNode[] = [];
		let afterParen = nameIndex + 1;
		const lparen = nameIndex >= 0
			? tokens.findIndex((t, idx) => idx > nameIndex && t.rawText === '(')
			: -1;
		if (lparen >= 0) {
			const parsed = this.parseParamList(tokens, lparen);
			params = parsed.params;
			afterParen = parsed.closeIndex + 1;
		}

		let returnType: string | undefined;
		if (isFunction && tokens[afterParen] && tokenWord(tokens[afterParen]) === 'as') {
			returnType = this.captureType(tokens, afterParen + 1);
		}
		return {
			kind: 'Declare',
			name: nameToken ? this.stripBrackets(nameToken.rawText) : '',
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
		const nameToken = group[i];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		i++;
		let isArray = false;
		if (group[i] && group[i].rawText === '(') {
			isArray = true;
			i = this.skipParens(group, i);
		}
		let isNew = false;
		let asType: string | undefined;
		if (group[i] && tokenWord(group[i]) === 'as') {
			i++;
			if (group[i] && tokenWord(group[i]) === 'new') {
				isNew = true;
				i++;
			}
			asType = this.captureType(group, i, isConst ? '=' : undefined);
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
			asType,
			defaultRaw,
			isArray,
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
			if (this.isModuleLevelStarter(stmt)) {
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
			visibility,
			fields,
			closed,
			span: { start: head.start, end: (endStmt ?? head).end },
		};
	}

	private parseTypeField(stmt: LogicalStatement, tokens: VbaToken[]): TypeFieldNode {
		const name = this.stripBrackets(tokens[0].rawText);
		let i = 1;
		let isArray = false;
		if (tokens[i] && tokens[i].rawText === '(') {
			isArray = true;
			i = this.skipParens(tokens, i);
		}
		let asType: string | undefined;
		if (tokens[i] && tokenWord(tokens[i]) === 'as') {
			asType = this.captureType(tokens, i + 1);
		}
		return {
			kind: 'TypeField',
			name,
			asType,
			isArray,
			span: { start: stmt.start, end: stmt.end },
		};
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
				members.push({
					kind: 'EnumMember',
					name: this.stripBrackets(mtokens[0].rawText),
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

		const nameToken = tokens[i];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		i++;

		let params: ParameterNode[] = [];
		let afterParen = i;
		if (tokens[i] && tokens[i].rawText === '(') {
			const parsed = this.parseParamList(tokens, i);
			params = parsed.params;
			afterParen = parsed.closeIndex + 1;
		}

		let returnType: string | undefined;
		if (tokens[afterParen] && tokenWord(tokens[afterParen]) === 'as') {
			returnType = this.captureType(tokens, afterParen + 1);
		}

		const expected = this.procCloser(procKind);
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
			// Recovery: a new module-level construct means the End was forgotten.
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
			modifiers,
			params,
			returnType,
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
		const nameToken = group[i];
		const name = nameToken ? this.stripBrackets(nameToken.rawText) : '';
		i++;
		let isArray = false;
		if (group[i] && group[i].rawText === '(') {
			isArray = true;
			i = this.skipParens(group, i);
		}
		let asType: string | undefined;
		if (group[i] && tokenWord(group[i]) === 'as') {
			i++;
			asType = this.captureType(group, i, '=');
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
		const tokens = codeTokens(stmt);
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
		return this.makeBlockNode(opener, body, closed, span);
	}

	private makeBlockNode(
		opener: 'if' | 'for' | 'foreach' | 'do' | 'while' | 'with' | 'select',
		body: BodyNode[],
		closed: boolean,
		span: Span,
	): BodyNode {
		switch (opener) {
			case 'if':
				return { kind: 'IfBlock', body, closed, span } satisfies IfBlockNode;
			case 'for':
				return { kind: 'ForBlock', each: false, body, closed, span } satisfies ForBlockNode;
			case 'foreach':
				return { kind: 'ForBlock', each: true, body, closed, span } satisfies ForBlockNode;
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
		const tokens = codeTokens(stmt);
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
		const tokens = codeTokens(stmt);
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
		const groups: VbaToken[][] = [];
		let current: VbaToken[] = [];
		let depth = 0;
		for (let i = from; i < to; i++) {
			const t = tokens[i];
			if (t.rawText === '(') {
				depth++;
			} else if (t.rawText === ')') {
				depth--;
			}
			if (depth === 0 && t.rawText === ',') {
				groups.push(current);
				current = [];
				continue;
			}
			current.push(t);
		}
		groups.push(current);
		return groups;
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

	/** Index just past the ')' matching the '(' at lparen. */
	private matchParen(tokens: VbaToken[], lparen: number): number {
		let depth = 0;
		for (let i = lparen; i < tokens.length; i++) {
			if (tokens[i].rawText === '(') {
				depth++;
			} else if (tokens[i].rawText === ')') {
				depth--;
				if (depth === 0) {
					return i;
				}
			}
		}
		return tokens.length - 1;
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

	private canonical(token: VbaToken | undefined): string {
		if (!token) {
			return '';
		}
		return token.canonicalText ?? token.rawText;
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
