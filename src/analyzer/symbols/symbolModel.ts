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
//   - 5.2.3.5 External Procedure Declarations (Declare)
//   - 5.3   Procedure declarations (Sub/Function/Property)
//   - 5.4   Local variable declarations (Dim/Static inside a procedure)

import type { Span } from '../parser/nodes';
import type { VbaDoc } from '../docs/docModel';

/**
 * The workbook-project role of a module, as reported by the host
 * (the workbook engine's `listModules`). Maps to MS-VBAL 4.2 module kinds plus the two
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
	| 'event'
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
	/** Raw fixed-length suffix after `As String *`, when present. */
	fixedLength?: string;
	/** Raw default/constant initializer text after `=` when source provides one. */
	defaultRaw?: string;
	/** True when a parameter is declared Optional. */
	optional?: boolean;
	/** True when a parameter is declared ParamArray. */
	paramArray?: boolean;
	/** True when a parameter is explicitly ByVal. */
	byVal?: boolean;
	/** True when a parameter is explicitly ByRef. */
	byRef?: boolean;
	/** True when the declaration is an array. */
	isArray?: boolean;
	/**
	 * Declared `As New`. VBA instantiates such a variable on ANY access - even
	 * after `Set x = Nothing` - so it can never be Nothing when a member is
	 * touched.
	 */
	isAutoInstantiated?: boolean;
	/** Raw array bounds text when an array declaration is fixed-size. */
	arrayBounds?: string;
	/** External Declare statements are Function or Sub callables. */
	declareKind?: 'Function' | 'Sub';
	/** True when an external Declare includes PtrSafe. */
	ptrSafe?: boolean;
	/** Library name from `Lib "..."` on an external Declare. */
	libName?: string;
	/** Alias name from `Alias "..."` on an external Declare. */
	aliasName?: string;
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
	defaultRaw?: string;
	byVal?: boolean;
	byRef?: boolean;
}

/** Exported callable signature collected from the project symbol graph. */
export interface VbaProcedureSignature {
	name: string;
	moduleName: string;
	kind: Extract<VbaSymbolKind, 'sub' | 'function'>;
	params: VbaProcedureParam[];
	returnType?: string;
	signature?: string;
	visibility?: SymbolVisibility;
	doc?: VbaDoc;
	external?: boolean;
	ptrSafe?: boolean;
	libName?: string;
	aliasName?: string;
}

export function procedureKindKeyword(kind: VbaProcedureSignature['kind']): string {
	return kind === 'function' ? 'Function' : 'Sub';
}

function quoteVbaString(value: string): string {
	return `"${value.replace(/"/g, '""')}"`;
}

export function formatProcedureParamLabel(
	param: VbaProcedureParam,
	options: { includePassing?: boolean } = {},
): string {
	let label = param.name;
	if (param.isArray) {
		label += '()';
	}
	if (param.type) {
		label += ` As ${param.type}`;
	}
	if (param.defaultRaw) {
		label += ` = ${param.defaultRaw}`;
	}
	if (param.paramArray) {
		label = `ParamArray ${label}`;
	}
	if (options.includePassing && !param.paramArray) {
		if (param.byVal) {
			label = `ByVal ${label}`;
		} else if (param.byRef) {
			label = `ByRef ${label}`;
		}
	}
	return param.optional ? `[${label}]` : label;
}

export function procedureSignatureLabel(procedure: Pick<
	VbaProcedureSignature,
	'name' | 'kind' | 'params' | 'returnType' | 'external'
>): string {
	const params = procedure.params
		.map((param) => formatProcedureParamLabel(param, { includePassing: procedure.external }))
		.join(', ');
	const returns = procedure.kind === 'function' && procedure.returnType
		? ` As ${procedure.returnType}`
		: '';
	return `${procedure.name}(${params})${returns}`;
}

export function procedureDeclarationSignature(
	procedure: Pick<
		VbaProcedureSignature,
		'name' | 'kind' | 'params' | 'returnType' | 'external' | 'ptrSafe' | 'libName' | 'aliasName'
	>,
): string {
	if (procedure.external) {
		const keyword = procedureKindKeyword(procedure.kind);
		const ptrSafe = procedure.ptrSafe ? 'PtrSafe ' : '';
		const lib = procedure.libName ? ` Lib ${quoteVbaString(procedure.libName)}` : '';
		const alias = procedure.aliasName ? ` Alias ${quoteVbaString(procedure.aliasName)}` : '';
		const externalTarget = lib || alias ? `${lib}${alias} ` : '';
		const params = procedure.params
			.map((param) => formatProcedureParamLabel(param, { includePassing: true }))
			.join(', ');
		const returns = procedure.kind === 'function' && procedure.returnType
			? ` As ${procedure.returnType}`
			: '';
		return `Declare ${ptrSafe}${keyword} ${procedure.name}${externalTarget}(${params})${returns}`;
	}
	return `${procedureKindKeyword(procedure.kind)} ${procedureSignatureLabel(procedure)}`;
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
	/** Span of the type identifier, or the module start for object modules. */
	nameSpan?: Span;
	/** Full declaration span, or the module span for object modules. */
	fullSpan?: Span;
	/** Visibility on Type/Enum declarations when present. */
	visibility?: SymbolVisibility;
	/** Inline `'''` documentation attached to the type/module declaration. */
	doc?: VbaDoc;
}

/** Source declaration location for a project object member. */
export interface VbaProjectClassMemberDefinition {
	moduleName: string;
	/** Span of the declared member identifier itself. */
	nameSpan: Span;
	/** Full span of the declaration. */
	fullSpan: Span;
}

/** A source-declared member of a project object module or user-defined Type. */
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

/** Public member surface for a workbook-defined object type, standard module, or user-defined Type. */
export interface VbaProjectClassMembers {
	name: string;
	kind: Extract<VbaProjectTypeKind, 'class' | 'document' | 'userform' | 'userType' | 'enum'> | 'standardModule';
	moduleName: string;
	/** Interfaces named by module-level Implements statements. */
	implements?: string[];
	/** Inline `'''` documentation attached to the object module header. */
	doc?: VbaDoc;
	/**
	 * True when the member list is complete enough to prove absence. Plain class
	 * modules are source-exhaustive; document modules and UserForms also expose
	 * host/designer members, so their source-only surface is not exhaustive yet.
	 */
	exhaustive?: boolean;
	/**
	 * Whether the module has a default instance - `Attribute VB_PredeclaredId
	 * = True`, which makes the module's own name usable as a value. Document
	 * modules and UserForms always have one; a plain class module does not,
	 * and using its bare name where a value belongs is `Variable not defined`
	 * (issue #47).
	 *
	 * Three states, not two: `true` and `false` are answers, and ABSENT means
	 * the attribute header was never read - no rule may assume either way,
	 * because the attribute is invisible in the code pane and a host that
	 * cannot see it would otherwise turn every predeclared class red.
	 */
	predeclaredId?: boolean;
	members: VbaProjectClassMember[];
}

/** Lowercased key used for module-qualified procedure lookups. */
export function qualifiedProcedureKey(moduleName: string, name: string): string {
	return `${moduleName.toLowerCase()}.${name.toLowerCase()}`;
}

/** True for bare callables: Sub, Function, and external Declare statements. */
export function isBareCallableKind(kind: VbaSymbolKind): boolean {
	return kind === 'sub' || kind === 'function' || kind === 'declare';
}

/** Converts a symbol's parameter children into the shared callable parameter model. */
export function procedureParamsFromSymbol(
	symbol: VbaSymbol,
	options: { includePassing?: boolean } = {},
): VbaProcedureParam[] {
	return (symbol.children ?? [])
		.filter((child) => child.kind === 'parameter')
		.map((child) => {
			const param: VbaProcedureParam = {
				name: child.name,
				type: child.asType,
				optional: child.optional ?? false,
				paramArray: child.paramArray ?? false,
				isArray: child.isArray ?? false,
			};
			if (child.defaultRaw !== undefined) {
				param.defaultRaw = child.defaultRaw;
			}
			if (options.includePassing) {
				if (child.byVal) {
					param.byVal = true;
				} else if (child.byRef) {
					param.byRef = true;
				}
			}
			return param;
		});
}

/** Converts a Sub/Function/Declare symbol into the shared callable signature model. */
export function procedureSignatureFromSymbol(
	symbol: VbaSymbol,
): VbaProcedureSignature | undefined {
	if (!isBareCallableKind(symbol.kind)) {
		return undefined;
	}
	const external = symbol.kind === 'declare';
	const kind = callableKindForSymbol(symbol);
	const signatureBase: VbaProcedureSignature = {
		name: symbol.name,
		moduleName: symbol.moduleName,
		kind,
		params: procedureParamsFromSymbol(symbol, { includePassing: true }),
		returnType: symbol.asType,
		visibility: symbol.visibility,
		doc: symbol.doc,
		external: external || undefined,
		ptrSafe: symbol.ptrSafe,
		libName: symbol.libName,
		aliasName: symbol.aliasName,
	};
	return {
		...signatureBase,
		signature: external
			? procedureDeclarationSignature(signatureBase)
			: procedureSignatureLabel(signatureBase),
	};
}

function callableKindForSymbol(symbol: VbaSymbol): Extract<VbaSymbolKind, 'sub' | 'function'> {
	if (symbol.kind === 'declare') {
		return symbol.declareKind === 'Function' ? 'function' : 'sub';
	}
	return symbol.kind as Extract<VbaSymbolKind, 'sub' | 'function'>;
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

/** True for the five source procedure body symbol kinds. */
export function isProcedureKind(kind: VbaSymbolKind): boolean {
	return (
		kind === 'sub' ||
		kind === 'function' ||
		kind === 'propertyGet' ||
		kind === 'propertyLet' ||
		kind === 'propertySet'
	);
}
