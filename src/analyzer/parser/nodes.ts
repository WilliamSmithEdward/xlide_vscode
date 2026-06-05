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
	visibility?: string;
	fields: TypeFieldNode[];
	closed: boolean;
}

export interface TypeFieldNode extends NodeBase {
	kind: 'TypeField';
	name: string;
	asType?: string;
	fixedLength?: string;
	isArray: boolean;
}

/** Enum ... End Enum (MS-VBAL 5.2.3.4). */
export interface EnumNode extends NodeBase {
	kind: 'Enum';
	name: string;
	visibility?: string;
	members: EnumMemberNode[];
	closed: boolean;
}

export interface EnumMemberNode extends NodeBase {
	kind: 'EnumMember';
	name: string;
	/** Raw member value expression after `=`, when present. */
	valueRaw?: string;
}

/** The kind of procedure (MS-VBAL 5.3). */
export type ProcKind = 'Sub' | 'Function' | 'PropertyGet' | 'PropertyLet' | 'PropertySet';

export interface ProcedureNode extends NodeBase {
	kind: 'Procedure';
	procKind: ProcKind;
	name: string;
	modifiers: string[];
	params: ParameterNode[];
	returnType?: string;
	body: BodyNode[];
	/** True if a matching End Sub/Function/Property was found. */
	closed: boolean;
}

export interface ParameterNode extends NodeBase {
	kind: 'Parameter';
	name: string;
	optional: boolean;
	byVal: boolean;
	byRef: boolean;
	paramArray: boolean;
	asType?: string;
	isArray: boolean;
	defaultRaw?: string;
}

/** Generic catch-all statement (assignment, call, Set, Exit, GoTo, label, ...). */
export interface StatementNode extends NodeBase {
	kind: 'Statement';
	/** Raw source text of the statement (without separators). */
	raw: string;
}

// Block statements (MS-VBAL 5.4). Each owns a body and records whether it was
// properly closed, which drives block-mismatch diagnostics.

export interface IfBlockNode extends NodeBase {
	kind: 'IfBlock';
	body: BodyNode[];
	closed: boolean;
}

export interface ForBlockNode extends NodeBase {
	kind: 'ForBlock';
	each: boolean;
	controlVariable?: string;
	controlVariableSpan?: Span;
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
	| StatementNode
	| ConditionalDirectiveNode
	| VariableGroupNode
	| IfBlockNode
	| ForBlockNode
	| DoBlockNode
	| WhileBlockNode
	| WithBlockNode
	| SelectBlockNode;

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
