// VBA symbol model (Phase 4: Project-Wide Symbol Graph).
//
// These types are a thin, host-agnostic projection of the parser AST
// (src/analyzer/parser/nodes.ts) into named declarations with scope and span
// information. They carry no `vscode` dependency so the symbol layer can be
// unit-tested under vitest and reused from any host.
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025):
//   - 4.2   Modules (procedural vs class module)
//   - 5.2.3 Module variable declarations / 5.2.4 Const declarations
//   - 5.2.3.3 User-defined Type / 5.2.3.4 Enum
//   - 5.3   Procedure declarations (Sub/Function/Property)
//   - 5.4   Local variable declarations (Dim/Static inside a procedure)

import type { Span } from '../parser/nodes';
import type { VbaDoc } from '../docs/docModel';

/**
 * The workbook-project role of a module, as reported by the host
 * (pyOpenVBA `listModules`). Maps to MS-VBAL 4.2 module kinds plus the two
 * host-specific document kinds (worksheet/ThisWorkbook documents and UserForms).
 */
export type ModuleSymbolKind = 'standard' | 'class' | 'document' | 'userform';

/** Every kind of named symbol the index can produce. */
export type VbaSymbolKind =
	| 'module'
	| 'sub'
	| 'function'
	| 'propertyGet'
	| 'propertyLet'
	| 'propertySet'
	| 'parameter'
	| 'localVariable'
	| 'moduleVariable'
	| 'constant'
	| 'enum'
	| 'enumMember'
	| 'type'
	| 'typeField'
	| 'declare';

/** Declaration visibility (MS-VBAL 5.2.3.1 / 5.3.1.1). */
export type SymbolVisibility =
	| 'Public'
	| 'Private'
	| 'Friend'
	| 'Global'
	| 'Dim'
	| 'Static';

/** A single named declaration discovered in a module. */
export interface VbaSymbol {
	name: string;
	kind: VbaSymbolKind;
	/** Span of the declared identifier itself (selection range for go-to-def). */
	nameSpan: Span;
	/** Full span of the declaration (range for document symbols). */
	fullSpan: Span;
	/** Owning module's name (VB component name). */
	moduleName: string;
	/** Enclosing procedure/type/enum name, when nested. */
	containerName?: string;
	visibility?: SymbolVisibility;
	/** Declared `As` type, when present. */
	asType?: string;
	/** Raw default/constant initializer text after `=` when source provides one. */
	defaultRaw?: string;
	/** True when a parameter is declared Optional. */
	optional?: boolean;
	/** True when a parameter is declared ParamArray. */
	paramArray?: boolean;
	/** True when a parameter is explicitly ByVal. */
	byVal?: boolean;
	/** True when a parameter is explicitly or implicitly ByRef. */
	byRef?: boolean;
	/** True when the declaration is an array. */
	isArray?: boolean;
	/** Nested symbols (procedure params/locals, enum members, UDT fields). */
	children?: VbaSymbol[];
	/** Inline `'''` XML documentation comment attached to the declaration. */
	doc?: VbaDoc;
	/** Exported VBA attribute lines attached to this symbol, when source provides them. */
	attributes?: VbaSymbolAttribute[];
}

/** Exported VBA attribute attached to a module or member declaration. */
export interface VbaSymbolAttribute {
	name: string;
	valueRaw: string;
	/** Member target before the dot in `Attribute Value.VB_UserMemId = 0`, if present. */
	targetName?: string;
	/** Span of the attribute name or dotted target/name. */
	nameSpan: Span;
	/** Full span of the attribute line. */
	fullSpan: Span;
}

/** Parameter shape used by project-wide callable signature diagnostics. */
export interface VbaProcedureParam {
	name: string;
	type?: string;
	optional: boolean;
	paramArray: boolean;
	isArray: boolean;
}

/** Exported callable signature collected from the project symbol graph. */
export interface VbaProcedureSignature {
	name: string;
	moduleName: string;
	kind: Extract<VbaSymbolKind, 'sub' | 'function'>;
	params: VbaProcedureParam[];
	returnType?: string;
}

/** Project-defined type-name categories visible in `As` type positions. */
export type VbaProjectTypeKind =
	| 'class'
	| 'document'
	| 'userform'
	| 'enum'
	| 'userType';

/** A project-defined type name visible to a module. */
export interface VbaProjectTypeName {
	name: string;
	kind: VbaProjectTypeKind;
	/** Module where the type name is declared or represented. */
	moduleName: string;
	/** Visibility on Type/Enum declarations when present. */
	visibility?: SymbolVisibility;
}

/** Source declaration location for a project object member. */
export interface VbaProjectClassMemberDefinition {
	moduleName: string;
	/** Span of the declared member identifier itself. */
	nameSpan: Span;
	/** Full span of the declaration. */
	fullSpan: Span;
}

/** A source-declared member of a project object module. */
export interface VbaProjectClassMember {
	name: string;
	kind: 'property' | 'method' | 'event';
	/** Declared return/object type when the source provides one. */
	returns?: string;
	/** Source-backed callable signature for methods and argument-taking properties. */
	signature?: string;
	/** True when source proves assignment to the member is allowed. */
	writable?: boolean;
	/** Declared value type accepted by assignment when source provides one. */
	writeType?: string;
	moduleName: string;
	visibility?: SymbolVisibility;
	/** Inline `'''` XML documentation comment attached to the member declaration. */
	doc?: VbaDoc;
	/** Source declaration locations for go-to-definition. */
	definitions?: VbaProjectClassMemberDefinition[];
	/** True when exported source marks this member as the VBA default member. */
	defaultMember?: boolean;
	/** Exported attribute lines attached to the member declaration. */
	attributes?: VbaSymbolAttribute[];
}

/** Public member surface for a workbook-defined object type. */
export interface VbaProjectClassMembers {
	name: string;
	kind: Extract<VbaProjectTypeKind, 'class' | 'document' | 'userform'>;
	moduleName: string;
	/**
	 * True when the member list is complete enough to prove absence. Plain class
	 * modules are source-exhaustive; document modules and UserForms also expose
	 * host/designer members, so their source-only surface is not exhaustive yet.
	 */
	exhaustive?: boolean;
	members: VbaProjectClassMember[];
}

/** Lowercased key used for module-qualified procedure lookups. */
export function qualifiedProcedureKey(moduleName: string, name: string): string {
	return `${moduleName.toLowerCase()}.${name.toLowerCase()}`;
}

/** The symbol view of a single module. */
export interface ModuleSymbols {
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	/** The module itself as a top-level symbol (with all members as children). */
	root: VbaSymbol;
	/** Flat list of every symbol in the module, including nested ones. */
	all: VbaSymbol[];
}

/** True for the five procedure symbol kinds. */
export function isProcedureKind(kind: VbaSymbolKind): boolean {
	return (
		kind === 'sub' ||
		kind === 'function' ||
		kind === 'propertyGet' ||
		kind === 'propertyLet' ||
		kind === 'propertySet'
	);
}
