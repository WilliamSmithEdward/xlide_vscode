import type { ModuleNode } from './nodes';

// Parser snapshots are immutable. Keep the lexer-derived negative fact weakly
// so repeated symbol/semantic queries need not traverse every procedure body.
// Positive and caller-supplied trees retain the conditional AST traversal.
const directiveFreeModules = new WeakSet<ModuleNode>();

export function rememberDirectiveFreeModule(module: ModuleNode): void {
	directiveFreeModules.add(module);
}

export function isKnownDirectiveFreeModule(module: ModuleNode): boolean {
	return directiveFreeModules.has(module);
}
