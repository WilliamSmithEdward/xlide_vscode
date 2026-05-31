// Builds the symbol view of a single module from its parser AST (Phase 4).
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025):
//   - 5.2.3 / 5.2.4 module variable and Const declarations
//   - 5.2.3.3 Type ... End Type / 5.2.3.4 Enum ... End Enum
//   - 5.3 Sub/Function/Property procedures and parameter lists
//   - 5.4 local Dim/Static declarations (including block-nested)
//
// No `vscode` dependency: pure AST -> symbol projection, unit-tested directly.

import { tokenize } from '../lexer/tokenize';
import type {
	BodyNode,
	EnumNode,
	ModuleNode,
	ProcedureNode,
	ProcKind,
	Span,
	TypeNode,
	VariableGroupNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { extractLeadingDoc } from '../docs/docComment';
import type {
	ModuleSymbolKind,
	ModuleSymbols,
	SymbolVisibility,
	VbaSymbol,
	VbaSymbolKind,
} from './symbolModel';

/** Maps a procedure AST kind to its symbol kind. */
function procSymbolKind(procKind: ProcKind): VbaSymbolKind {
	switch (procKind) {
		case 'Sub':
			return 'sub';
		case 'Function':
			return 'function';
		case 'PropertyGet':
			return 'propertyGet';
		case 'PropertyLet':
			return 'propertyLet';
		case 'PropertySet':
			return 'propertySet';
	}
}

/** Normalizes a raw modifier word to a SymbolVisibility, if it is one. */
function toVisibility(word: string | undefined): SymbolVisibility | undefined {
	switch ((word ?? '').toLowerCase()) {
		case 'public':
			return 'Public';
		case 'private':
			return 'Private';
		case 'friend':
			return 'Friend';
		case 'global':
			return 'Global';
		case 'dim':
			return 'Dim';
		case 'static':
			return 'Static';
		default:
			return undefined;
	}
}

/** Picks the visibility keyword out of a procedure's modifier list. */
function procVisibility(modifiers: string[]): SymbolVisibility | undefined {
	for (const m of modifiers) {
		const v = toVisibility(m);
		if (v && v !== 'Dim' && v !== 'Static') {
			return v;
		}
	}
	return undefined;
}

/**
 * Locates the span of the first occurrence of `name` (case-insensitive identifier
 * or keyword token) inside `span`, using the real lexer so the result is
 * deterministic and never matches text inside a comment or string. Falls back to
 * the whole span when the token is not found.
 */
function locateName(source: string, span: Span, name: string): Span {
	const slice = source.slice(span.start, span.end);
	const lower = name.toLowerCase();
	for (const tok of tokenize(slice)) {
		if (
			(tok.kind === 'identifier' || tok.kind === 'keyword') &&
			tok.rawText.toLowerCase() === lower
		) {
			return { start: span.start + tok.start, end: span.start + tok.end };
		}
		if (tok.kind === 'bracketedIdentifier') {
			const inner = tok.rawText.replace(/^\[|\]$/g, '');
			if (inner.toLowerCase() === lower) {
				return { start: span.start + tok.start, end: span.start + tok.end };
			}
		}
	}
	return span;
}

/** Recursively collects local variable declarations inside a procedure body. */
function collectLocals(
	body: BodyNode[],
	source: string,
	moduleName: string,
	containerName: string,
	out: VbaSymbol[],
): void {
	for (const node of body) {
		if (node.kind === 'VariableGroup') {
			for (const decl of node.declarations) {
				out.push({
					name: decl.name,
					kind: node.isConst ? 'constant' : 'localVariable',
					nameSpan: locateName(source, decl.span, decl.name),
					fullSpan: decl.span,
					moduleName,
					containerName,
					visibility: toVisibility(node.modifier),
					asType: decl.asType,
				});
			}
		} else if ('body' in node && Array.isArray(node.body)) {
			collectLocals(node.body, source, moduleName, containerName, out);
		}
	}
}

/** Builds the symbol for one procedure, with parameter and local children. */
function buildProcedure(
	proc: ProcedureNode,
	source: string,
	moduleName: string,
	flat: VbaSymbol[],
): VbaSymbol {
	const children: VbaSymbol[] = [];
	const symbol: VbaSymbol = {
		name: proc.name,
		kind: procSymbolKind(proc.procKind),
		nameSpan: locateName(source, proc.span, proc.name),
		fullSpan: proc.span,
		moduleName,
		visibility: procVisibility(proc.modifiers),
		asType: proc.returnType,
		children,
	};

	for (const param of proc.params) {
		const paramSymbol: VbaSymbol = {
			name: param.name,
			kind: 'parameter',
			nameSpan: locateName(source, param.span, param.name),
			fullSpan: param.span,
			moduleName,
			containerName: proc.name,
			asType: param.asType,
		};
		children.push(paramSymbol);
		flat.push(paramSymbol);
	}

	const locals: VbaSymbol[] = [];
	collectLocals(proc.body, source, moduleName, proc.name, locals);
	for (const local of locals) {
		children.push(local);
		flat.push(local);
	}

	return symbol;
}

/** Builds the symbol for a Type ... End Type, with field children. */
function buildType(
	node: TypeNode,
	source: string,
	moduleName: string,
	flat: VbaSymbol[],
): VbaSymbol {
	const children: VbaSymbol[] = [];
	const symbol: VbaSymbol = {
		name: node.name,
		kind: 'type',
		nameSpan: locateName(source, node.span, node.name),
		fullSpan: node.span,
		moduleName,
		visibility: toVisibility(node.visibility),
		children,
	};
	for (const field of node.fields) {
		const fieldSymbol: VbaSymbol = {
			name: field.name,
			kind: 'typeField',
			nameSpan: locateName(source, field.span, field.name),
			fullSpan: field.span,
			moduleName,
			containerName: node.name,
			asType: field.asType,
		};
		children.push(fieldSymbol);
		flat.push(fieldSymbol);
	}
	return symbol;
}

/** Builds the symbol for an Enum ... End Enum, with member children. */
function buildEnum(
	node: EnumNode,
	source: string,
	moduleName: string,
	flat: VbaSymbol[],
): VbaSymbol {
	const children: VbaSymbol[] = [];
	const symbol: VbaSymbol = {
		name: node.name,
		kind: 'enum',
		nameSpan: locateName(source, node.span, node.name),
		fullSpan: node.span,
		moduleName,
		visibility: toVisibility(node.visibility),
		children,
	};
	for (const member of node.members) {
		const memberSymbol: VbaSymbol = {
			name: member.name,
			kind: 'enumMember',
			nameSpan: locateName(source, member.span, member.name),
			fullSpan: member.span,
			moduleName,
			containerName: node.name,
		};
		children.push(memberSymbol);
		flat.push(memberSymbol);
	}
	return symbol;
}

/** Adds module-level variable/const group declarations as flat + child symbols. */
function buildModuleVariables(
	group: VariableGroupNode,
	source: string,
	moduleName: string,
	rootChildren: VbaSymbol[],
	flat: VbaSymbol[],
): void {
	const doc = extractLeadingDoc(source, group.span.start);
	for (const decl of group.declarations) {
		const symbol: VbaSymbol = {
			name: decl.name,
			kind: group.isConst ? 'constant' : 'moduleVariable',
			nameSpan: locateName(source, decl.span, decl.name),
			fullSpan: decl.span,
			moduleName,
			visibility: toVisibility(group.modifier),
			asType: decl.asType,
			doc,
		};
		rootChildren.push(symbol);
		flat.push(symbol);
	}
}

/**
 * Builds the {@link ModuleSymbols} view of a module from its source text.
 *
 * @param moduleName VB component name (used as the module symbol name and as the
 *   owner of every contained symbol).
 * @param moduleKind Workbook-project role of the module (standard/class/etc.).
 * @param source Full VBA source of the module.
 */
export function buildModuleSymbols(
	moduleName: string,
	moduleKind: ModuleSymbolKind,
	source: string,
): ModuleSymbols {
	const module: ModuleNode = parseModule(source);
	const rootChildren: VbaSymbol[] = [];
	const flat: VbaSymbol[] = [];

	for (const member of module.members) {
		switch (member.kind) {
			case 'Procedure': {
				const proc = buildProcedure(member, source, moduleName, flat);
				proc.doc = extractLeadingDoc(source, member.span.start);
				rootChildren.push(proc);
				flat.push(proc);
				break;
			}
			case 'Type': {
				const type = buildType(member, source, moduleName, flat);
				type.doc = extractLeadingDoc(source, member.span.start);
				rootChildren.push(type);
				flat.push(type);
				break;
			}
			case 'Enum': {
				const en = buildEnum(member, source, moduleName, flat);
				en.doc = extractLeadingDoc(source, member.span.start);
				rootChildren.push(en);
				flat.push(en);
				break;
			}
			case 'VariableGroup': {
				buildModuleVariables(member, source, moduleName, rootChildren, flat);
				break;
			}
			case 'Declare': {
				const symbol: VbaSymbol = {
					name: member.name,
					kind: 'declare',
					nameSpan: locateName(source, member.span, member.name),
					fullSpan: member.span,
					moduleName,
					visibility: toVisibility(member.visibility),
					doc: extractLeadingDoc(source, member.span.start),
				};
				rootChildren.push(symbol);
				flat.push(symbol);
				break;
			}
			default:
				break;
		}
	}

	const moduleSymbolKind: VbaSymbolKind = 'module';
	const root: VbaSymbol = {
		name: moduleName,
		kind: moduleSymbolKind,
		nameSpan: { start: module.span.start, end: module.span.start },
		fullSpan: module.span,
		moduleName,
		children: rootChildren,
	};

	return { moduleName, moduleKind, root, all: flat };
}
