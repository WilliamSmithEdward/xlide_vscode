// VBA abstract syntax tree node definitions.
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025):
//   - 4.2   Modules (module header / module body)
//   - 5.1   Procedural Module Header / 5.2 Module Body Structure
//   - 5.2.1 Option Directives
//   - 5.2.3 Module Variable Declarations
//   - 5.2.4 Const Declarations
//   - 3.4   Conditional Compilation Directives
//   - 5.2.2 Implicit / 5.3 Procedure declarations (Sub/Function/Property)
//   - 5.2.3.3 User Defined Type (Type ... End Type)
//   - 5.2.3.4 Enum Declarations (Enum ... End Enum)
//   - 5.4   Statements (block statements: If/For/Do/While/With/Select)
//   - 5.4.2 Call statements / 5.4.3 Assignment statements
//   - 5.6   Value expressions (ExprNode hierarchy - partial)
//
// The parser is error-tolerant: every node carries an absolute source span so a
// VS Code provider can map it to a Position with TextDocument.positionAt, and
// malformed input still yields a best-effort tree (Phase 3 acceptance criteria).

/** Absolute UTF-16 source offsets [start, end). */
export interface Span {
	start: number;
	end: number;
}

/** Severity of a parse-time diagnostic. */
export type ParseSeverity = 'error' | 'warning';

/** A diagnostic produced while parsing (block mismatches, etc.). */
export interface ParseDiagnostic {
	span: Span;
	message: string;
	severity: ParseSeverity;
	/** MS-VBAL section that justifies the rule, when applicable. */
	specRef?: string;
}

/** Discriminated-union tag for every node. */
export type NodeKind =
	| 'Module'
	| 'Attribute'
	| 'Option'
	| 'ConditionalDirective'
	| 'Declare'
	| 'Event'
	| 'VariableGroup'
	| 'VariableDecl'
	| 'Type'
	| 'TypeField'
	| 'Enum'
	| 'EnumMember'
	| 'Procedure'
	| 'Parameter'
	| 'Assignment'
	| 'Call'
	| 'IfBlock'
	| 'ForBlock'
	| 'DoBlock'
	| 'WhileBlock'
	| 'WithBlock'
	| 'SelectBlock'
	| 'Statement';

export interface NodeBase {
	kind: NodeKind;
	span: Span;
}

/** What kind of module this is (MS-VBAL 4.2). */
export type ModuleKind = 'procedural' | 'class' | 'unknown';

export interface ModuleNode extends NodeBase {
	kind: 'Module';
	moduleKind: ModuleKind;
	members: ModuleMember[];
	diagnostics: ParseDiagnostic[];
}

/** Attribute name/value line, e.g. Attribute VB_Name = "Module1" (MS-VBAL 4.2). */
export interface AttributeNode extends NodeBase {
	kind: 'Attribute';
	/** Raw attribute target/name text after `Attribute`, e.g. `VB_Name` or `Value.VB_UserMemId`. */
	name: string;
	/** Span of the raw attribute target/name text. */
	nameSpan: Span;
	/** Raw value text (right of '='), unparsed. */
	valueRaw: string;
}

/** Option directive (MS-VBAL 5.2.1): Explicit / Base n / Compare m / Private Module. */
export interface OptionNode extends NodeBase {
	kind: 'Option';
	/** Canonical text after "Option", e.g. "Explicit", "Base 1", "Compare Text". */
	optionText: string;
}

/** Conditional-compilation directive kind (MS-VBAL 3.4). */
export type ConditionalDirectiveKind =
	| 'Const'
	| 'If'
	| 'ElseIf'
	| 'Else'
	| 'EndIf'
	| 'Unknown';

/** `#Const`, `#If`, `#ElseIf`, `#Else`, or `#End If` directive (MS-VBAL 3.4). */
export interface ConditionalDirectiveNode extends NodeBase {
	kind: 'ConditionalDirective';
	directiveKind: ConditionalDirectiveKind;
	/** `#Const` compiler constant name. */
	name?: string;
	nameSpan?: Span;
	/** Raw `#Const` value expression, preserving source spacing. */
	valueRaw?: string;
	valueSpan?: Span;
	/** Raw `#If` / `#ElseIf` condition expression, without the trailing Then. */
	conditionRaw?: string;
	conditionSpan?: Span;
}

/** Declare statement for an external procedure (MS-VBAL 5.2.3.5). */
export interface DeclareNode extends NodeBase {
	kind: 'Declare';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	typeSuffix?: string;
	typeSuffixSpan?: Span;
	hasAsClause?: boolean;
	isFunction: boolean;
	visibility?: string;
	ptrSafe: boolean;
	libName?: string;
	aliasName?: string;
	params: ParameterNode[];
	returnType?: string;
}

/** Event declaration in a class/document/UserForm module (MS-VBAL 5.2.5). */
export interface EventNode extends NodeBase {
	kind: 'Event';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	visibility?: string;
	params: ParameterNode[];
}

/** Access modifier on a declaration/procedure. */
export type Visibility = 'Public' | 'Private' | 'Friend' | 'Global' | 'Dim' | 'Static';

/** A group of variable declarations sharing one modifier (MS-VBAL 5.2.3 / 5.2.4). */
export interface VariableGroupNode extends NodeBase {
	kind: 'VariableGroup';
	modifier: string;
	isConst: boolean;
	withEvents: boolean;
	declarations: VariableDeclNode[];
}

/** A single declared name within a VariableGroup. */
export interface VariableDeclNode extends NodeBase {
	kind: 'VariableDecl';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	typeSuffix?: string;
	typeSuffixSpan?: Span;
	hasAsClause?: boolean;
	asType?: string;
	fixedLength?: string;
	defaultRaw?: string;
	isArray: boolean;
	arrayBounds?: string;
	isNew: boolean;
}

/** User-defined Type ... End Type (MS-VBAL 5.2.3.3). */
export interface TypeNode extends NodeBase {
	kind: 'Type';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	visibility?: string;
	fields: TypeFieldNode[];
	/** Conditional-compilation directives (`#If` etc.) inside the Type body. */
	directives?: ConditionalDirectiveNode[];
	closed: boolean;
}

export interface TypeFieldNode extends NodeBase {
	kind: 'TypeField';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	typeSuffix?: string;
	typeSuffixSpan?: Span;
	hasAsClause?: boolean;
	asType?: string;
	fixedLength?: string;
	isArray: boolean;
}

/** Enum ... End Enum (MS-VBAL 5.2.3.4). */
export interface EnumNode extends NodeBase {
	kind: 'Enum';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	visibility?: string;
	members: EnumMemberNode[];
	/** Conditional-compilation directives (`#If` etc.) inside the Enum body. */
	directives?: ConditionalDirectiveNode[];
	closed: boolean;
}

export interface EnumMemberNode extends NodeBase {
	kind: 'EnumMember';
	name: string;
	/** Span of the declared name token. */
	nameSpan?: Span;
	/** Raw member value expression after `=`, when present. */
	valueRaw?: string;
}

/** The kind of procedure (MS-VBAL 5.3). */
export type ProcKind = 'Sub' | 'Function' | 'PropertyGet' | 'PropertyLet' | 'PropertySet';

export interface ProcedureNode extends NodeBase {
	kind: 'Procedure';
	procKind: ProcKind;
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	typeSuffix?: string;
	typeSuffixSpan?: Span;
	hasAsClause?: boolean;
	modifiers: string[];
	params: ParameterNode[];
	returnType?: string;
	/** Exported member metadata lines such as `Attribute Value.VB_UserMemId = 0`. */
	attributes?: AttributeNode[];
	body: BodyNode[];
	/** True if a matching End Sub/Function/Property was found. */
	closed: boolean;
}

export interface ParameterNode extends NodeBase {
	kind: 'Parameter';
	name: string;
	/** Span of the declared name token (absent when the name is missing). */
	nameSpan?: Span;
	typeSuffix?: string;
	typeSuffixSpan?: Span;
	hasAsClause?: boolean;
	optional: boolean;
	byVal: boolean;
	byRef: boolean;
	paramArray: boolean;
	asType?: string;
	isArray: boolean;
	defaultRaw?: string;
}

// ---------------------------------------------------------------------------
// Expressions (MS-VBAL §5.6 - partial; see MS-VBAL.verification-map.md)
// ---------------------------------------------------------------------------

/** Discriminated-union tag for expression nodes. */
export type ExprKind =
	| 'LiteralExpr'
	| 'IdentifierExpr'
	| 'MemberAccessExpr'
	| 'IndexExpr'
	| 'ParenExpr'
	| 'UnaryExpr'
	| 'BinaryExpr'
	| 'NewExpr'
	| 'AddressOfExpr'
	| 'TypeOfIsExpr';

export interface ExprBase {
	exprKind: ExprKind;
	span: Span;
}

/** Literal value kind per MS-VBAL §5.6 literal-expression. */
export type LiteralKind =
	| 'integer'
	| 'float'
	| 'string'
	| 'date'
	| 'boolean'
	| 'nothing'
	| 'null'
	| 'empty';

/** A literal value: number, string, date, boolean, or Nothing/Null/Empty. */
export interface LiteralExpr extends ExprBase {
	exprKind: 'LiteralExpr';
	literalKind: LiteralKind;
	raw: string;
}

/** A simple identifier reference - local, module-level, or project-visible name. */
export interface IdentifierExpr extends ExprBase {
	exprKind: 'IdentifierExpr';
	name: string;
	/** Type-declaration character (%, &, #, !, @, $) if present. */
	typeSuffix?: string;
}

/**
 * `object.member` member-access expression, or leading `.member` inside a
 * With block. When `object` is null the member resolves against the innermost
 * With receiver.
 */
export interface MemberAccessExpr extends ExprBase {
	exprKind: 'MemberAccessExpr';
	object: ExprNode | null;
	member: string;
	memberSpan: Span;
	/**
	 * `'bang'` for the `receiver!name` form (sugar for the receiver's default
	 * member indexed by the string `"name"`, e.g. `rs!Field` ≈ `rs.Fields("Field")`).
	 * Omitted / `'dot'` for ordinary `.member` access. Bang names are NOT literal
	 * members of the receiver type, so member-existence rules must treat them
	 * differently from dot access.
	 */
	accessKind?: 'dot' | 'bang';
}

/**
 * One entry in a call / index argument list (MS-VBAL §5.6.9 argument-list and
 * §5.4.2 call-statement argument-list). VBA arguments may be positional, named
 * (`name:=value`), or omitted (`Foo(a, , c)` - an empty slot that tells the
 * callee to use the parameter's default).
 */
export interface Argument {
	/** Named-argument name (without the `:=`), or undefined for a positional arg. */
	name?: string;
	/** Span of the name token, when `name` is present. */
	nameSpan?: Span;
	/** The argument value expression, or null for an omitted slot. */
	value: ExprNode | null;
	/** Span of the whole argument entry (name + value, or just the value). */
	span: Span;
}

/** `callee(arg, ...)` - covers function calls and array indexing. */
export interface IndexExpr extends ExprBase {
	exprKind: 'IndexExpr';
	callee: ExprNode;
	args: Argument[];
}

/**
 * `(inner)` parenthesized expression.
 * Wrapping a ByRef argument in parens coerces it to ByVal at the call site.
 */
export interface ParenExpr extends ExprBase {
	exprKind: 'ParenExpr';
	inner: ExprNode;
}

export type UnaryOperator = '-' | '+' | 'Not';

export type BinaryOperator =
	| '+' | '-' | '*' | '/' | '\\' | 'Mod' | '^'
	| '&'
	| '=' | '<>' | '<' | '>' | '<=' | '>='
	| 'And' | 'Or' | 'Xor' | 'Eqv' | 'Imp'
	| 'Like' | 'Is';

export interface UnaryExpr extends ExprBase {
	exprKind: 'UnaryExpr';
	operator: UnaryOperator;
	operand: ExprNode;
}

export interface BinaryExpr extends ExprBase {
	exprKind: 'BinaryExpr';
	operator: BinaryOperator;
	left: ExprNode;
	right: ExprNode;
}

/** `New TypeName` expression. */
export interface NewExpr extends ExprBase {
	exprKind: 'NewExpr';
	typeName: string;
	typeNameSpan: Span;
}

/** `AddressOf procedureName` expression. */
export interface AddressOfExpr extends ExprBase {
	exprKind: 'AddressOfExpr';
	target: IdentifierExpr;
}

/** `TypeOf expr Is TypeName` expression. */
export interface TypeOfIsExpr extends ExprBase {
	exprKind: 'TypeOfIsExpr';
	operand: ExprNode;
	typeName: string;
	typeNameSpan: Span;
}

/** Any expression node (MS-VBAL §5.6). */
export type ExprNode =
	| LiteralExpr
	| IdentifierExpr
	| MemberAccessExpr
	| IndexExpr
	| ParenExpr
	| UnaryExpr
	| BinaryExpr
	| NewExpr
	| AddressOfExpr
	| TypeOfIsExpr;

// ---------------------------------------------------------------------------
// Structured statement nodes (MS-VBAL §5.4)
// ---------------------------------------------------------------------------

/**
 * `[Let] lhs = rhs` or `Set lhs = rhs` - MS-VBAL §5.4.3.
 * `isSet` distinguishes object-reference assignment from value assignment.
 */
export interface AssignmentNode extends NodeBase {
	kind: 'Assignment';
	isSet: boolean;
	isLet: boolean;
	lhs: ExprNode;
	rhs: ExprNode;
}

/**
 * `[Call] callee [(args)]` or implicit call `callee arg, ...` - MS-VBAL §5.4.2.
 * When `hasCallKeyword` is true the argument list must be parenthesised.
 */
export interface CallNode extends NodeBase {
	kind: 'Call';
	hasCallKeyword: boolean;
	callee: ExprNode;
	args: Argument[];
}

/** Generic catch-all statement (Exit, GoTo, label, Return, and anything not yet
 *  parsed into a structured node). */
export interface StatementNode extends NodeBase {
	kind: 'Statement';
	/** Raw source text of the statement (without separators). */
	raw: string;
}

// Block statements (MS-VBAL 5.4). Each owns a body and records whether it was
// properly closed, which drives block-mismatch diagnostics.

/**
 * One arm of an `If` block (MS-VBAL 5.4.2.1): the leading `If`, zero or more
 * `ElseIf`, and an optional `Else`. Branch modeling makes flow-sensitive
 * analysis possible (each arm's statements and entry condition are explicit)
 * without disturbing `IfBlockNode.body`, which stays a flat list for the generic
 * body walkers.
 */
export interface IfBranchNode {
	kind: 'IfBranch';
	branchKind: 'if' | 'elseif' | 'else';
	/** Entry condition for `if`/`elseif` arms; null for `else` (or an unparsable condition). */
	condition: ExprNode | null;
	/** Raw condition text between the keyword and `Then` (absent for `else`). */
	conditionRaw?: string;
	conditionSpan?: Span;
	/** Statements in this arm only (excludes the arm's own header line). */
	body: BodyNode[];
	/** Span of the arm header line (`If ... Then` / `ElseIf ... Then` / `Else`). */
	headerSpan: Span;
	span: Span;
}

export interface IfBlockNode extends NodeBase {
	kind: 'IfBlock';
	/** Structured arms: always begins with the `if` arm, then any `elseif`/`else`. */
	branches: IfBranchNode[];
	/**
	 * Flat list of every arm's statements (and the `ElseIf`/`Else` header lines
	 * as raw `Statement`s), in source order. Retained so generic body walkers and
	 * block-balance diagnostics are unaffected by branch modeling.
	 */
	body: BodyNode[];
	closed: boolean;
}

export interface ForBlockNode extends NodeBase {
	kind: 'ForBlock';
	each: boolean;
	controlVariable?: string;
	controlVariableSpan?: Span;
	sourceExpression?: string;
	sourceExpressionSpan?: Span;
	nextVariable?: string;
	nextVariableSpan?: Span;
	body: BodyNode[];
	closed: boolean;
}

export interface DoBlockNode extends NodeBase {
	kind: 'DoBlock';
	body: BodyNode[];
	closed: boolean;
}

export interface WhileBlockNode extends NodeBase {
	kind: 'WhileBlock';
	body: BodyNode[];
	closed: boolean;
}

export interface WithBlockNode extends NodeBase {
	kind: 'WithBlock';
	body: BodyNode[];
	closed: boolean;
}

export interface SelectBlockNode extends NodeBase {
	kind: 'SelectBlock';
	body: BodyNode[];
	closed: boolean;
}

/** Any node that can appear inside a procedure body. */
export type BodyNode =
	| AssignmentNode
	| CallNode
	| StatementNode
	| ConditionalDirectiveNode
	| VariableGroupNode
	| IfBlockNode
	| ForBlockNode
	| DoBlockNode
	| WhileBlockNode
	| WithBlockNode
	| SelectBlockNode;

/**
 * A "leaf" executable statement: a structured assignment or call, or the raw
 * `Statement` catch-all. These are the body nodes that carry a single statement
 * span (no nested `body`); traversals that visit straight-line statements treat
 * all three uniformly via `isLeafStatement`. Every consumer reads only `.span`
 * (re-tokenizing the source), so `Assignment`/`Call` are span-compatible with
 * the raw `Statement` they replace.
 */
export type LeafStatementNode = AssignmentNode | CallNode | StatementNode;

/** True for the leaf statement nodes (Assignment / Call / raw Statement). */
export function isLeafStatement(node: BodyNode): node is LeafStatementNode {
	return node.kind === 'Assignment' || node.kind === 'Call' || node.kind === 'Statement';
}

/** Any node that can appear at module level. */
export type ModuleMember =
	| AttributeNode
	| OptionNode
	| ConditionalDirectiveNode
	| DeclareNode
	| EventNode
	| VariableGroupNode
	| TypeNode
	| EnumNode
	| ProcedureNode
	| StatementNode;
