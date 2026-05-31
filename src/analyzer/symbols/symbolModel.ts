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
	/** Nested symbols (procedure params/locals, enum members, UDT fields). */
	children?: VbaSymbol[];
	/** Inline `'''` XML documentation comment attached to the declaration. */
	doc?: VbaDoc;
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
