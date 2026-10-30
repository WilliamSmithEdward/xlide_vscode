import { conditionalCompilationKey, type ConditionalCompilationEnvironment } from '../conditional/conditionalCompilation';
import { buildModuleSymbols } from './buildModuleSymbols';
import type { ModuleSymbolKind, ModuleSymbols } from './symbolModel';

// Shared only by read-only editor resolvers: other symbol-builder consumers
// may mutate their graphs. Completion and hover results and external facts
// are constructed anew; callers must never mutate the borrowed symbol graph.
// Eight entries match the parser's bounded source cache; source edits, module
// renames, module-kind changes and another conditional compilation (32-bit
// Office's Win64, a project's arguments) all select a different snapshot.
const SYMBOL_CACHE_LIMIT = 8;
const symbolSnapshots: {
	source: string;
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	environment: string;
	symbols: ModuleSymbols;
}[] = [];

export function editorModuleSymbols(
	moduleName: string,
	moduleKind: ModuleSymbolKind,
	source: string,
	conditionalCompilation?: ConditionalCompilationEnvironment,
): ModuleSymbols {
	const environment = conditionalCompilationKey(conditionalCompilation);
	const index = symbolSnapshots.findIndex(entry =>
		entry.source === source && entry.moduleName === moduleName && entry.moduleKind === moduleKind
		&& entry.environment === environment);
	if (index >= 0) {
		const [entry] = symbolSnapshots.splice(index, 1);
		entry.source = source;
		symbolSnapshots.push(entry);
		return entry.symbols;
	}
	const symbols = buildModuleSymbols(moduleName, moduleKind, source, conditionalCompilation ? { conditionalCompilation } : {});
	symbolSnapshots.push({ source, moduleName, moduleKind, environment, symbols });
	if (symbolSnapshots.length > SYMBOL_CACHE_LIMIT) { symbolSnapshots.shift(); }
	return symbols;
}

