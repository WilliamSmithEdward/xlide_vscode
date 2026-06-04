import {
	ProjectIndex,
	type ModuleSymbolKind,
	type Span,
	type VbaSymbol,
	type VbaSymbolKind,
} from './analyzer';

export interface VbaPresentedSymbol {
	name: string;
	kind: VbaSymbolKind;
	detail: string;
	moduleName: string;
	containerName?: string;
	nameSpan: Span;
	fullSpan: Span;
	children: VbaPresentedSymbol[];
}

export interface VbaPresentedWorkspaceSymbol {
	name: string;
	kind: VbaSymbolKind;
	detail: string;
	moduleName: string;
	containerName: string;
	nameSpan: Span;
	fullSpan: Span;
}

export function symbolDetail(symbol: Pick<VbaSymbol, 'kind' | 'asType' | 'declareKind'>): string {
	switch (symbol.kind) {
		case 'module':
			return 'Module';
		case 'sub':
			return 'Sub';
		case 'function':
			return symbol.asType ? `Function As ${symbol.asType}` : 'Function';
		case 'propertyGet':
			return symbol.asType ? `Property Get As ${symbol.asType}` : 'Property Get';
		case 'propertyLet':
			return 'Property Let';
		case 'propertySet':
			return 'Property Set';
		case 'parameter':
			return symbol.asType ? `Parameter As ${symbol.asType}` : 'Parameter';
		case 'localVariable':
			return symbol.asType ? `Local As ${symbol.asType}` : 'Local';
		case 'moduleVariable':
			return symbol.asType ? `Variable As ${symbol.asType}` : 'Variable';
		case 'constant':
			return symbol.asType ? `Const As ${symbol.asType}` : 'Const';
		case 'enum':
			return 'Enum';
		case 'enumMember':
			return 'Enum member';
		case 'type':
			return 'Type';
		case 'typeField':
			return symbol.asType ? `Field As ${symbol.asType}` : 'Field';
		case 'event':
			return 'Event';
		case 'declare':
			return symbol.declareKind === 'Sub'
				? 'Declare Sub'
				: symbol.asType
					? `Declare Function As ${symbol.asType}`
					: 'Declare Function';
	}
}

export function documentOutlineSymbols(
	project: ProjectIndex,
	moduleName: string,
): VbaPresentedSymbol[] {
	return (project.documentSymbols(moduleName)?.children ?? []).map(presentDocumentSymbol);
}

export function documentOutlineSymbolsForSource(
	moduleName: string,
	moduleKind: ModuleSymbolKind,
	source: string,
): VbaPresentedSymbol[] {
	const project = new ProjectIndex();
	project.setModule({ moduleName, moduleKind, source });
	return documentOutlineSymbols(project, moduleName);
}

export function workspaceSymbols(
	project: ProjectIndex,
	query?: string,
): VbaPresentedWorkspaceSymbol[] {
	return project.workspaceSymbols(query).map((symbol) => ({
		name: symbol.name,
		kind: symbol.kind,
		detail: symbolDetail(symbol),
		moduleName: symbol.moduleName,
		containerName: symbol.containerName ?? symbol.moduleName,
		nameSpan: symbol.nameSpan,
		fullSpan: symbol.fullSpan,
	}));
}

function presentDocumentSymbol(symbol: VbaSymbol): VbaPresentedSymbol {
	return {
		name: symbol.name,
		kind: symbol.kind,
		detail: symbolDetail(symbol),
		moduleName: symbol.moduleName,
		containerName: symbol.containerName,
		nameSpan: symbol.nameSpan,
		fullSpan: symbol.fullSpan,
		children: (symbol.children ?? []).map(presentDocumentSymbol),
	};
}
