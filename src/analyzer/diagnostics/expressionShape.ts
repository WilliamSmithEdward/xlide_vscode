// Conservative declaration shapes for operands. Resolve only declarations we
// own; Object, Variant, unresolved members and ambiguous call results stay unknown.
import type { ExprNode, ProcedureNode } from '../parser/nodes';
import type { VbaSymbol } from '../symbols/symbolModel';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import { declarationShapeEnvironmentFor, typeEnvironmentFor, sourceNameScopeFor, type SourceNameScope, type DeclaredValueShape } from './typeInference';

type ShapeFactory = (proc: ProcedureNode) => (expr: ExprNode) => DeclaredValueShape | undefined;
const factories = new WeakMap<ReturnType<typeof buildModuleSymbols>, { visible: readonly VbaSymbol[] | undefined; factory: ShapeFactory }>();

export function expressionShapeResolver(symbols: ReturnType<typeof buildModuleSymbols>, projectVisibleSymbols?: readonly VbaSymbol[]): ShapeFactory {
	const cached = factories.get(symbols);
	if (cached && cached.visible === projectVisibleSymbols) { return cached.factory; }
	const declarations = new Map<string, NonNullable<typeof symbols.root.children>>();
	for (const symbol of symbols.root.children ?? []) {
		const key = symbol.name.toLowerCase();
		const group = declarations.get(key) ?? [];
		group.push(symbol);
		declarations.set(key, group);
	}
	const fields = new WeakMap<VbaSymbol, ReadonlyMap<string, VbaSymbol>>();
	const procedures = new WeakMap<ProcedureNode, ReturnType<ShapeFactory>>();
	const factory: ShapeFactory = (proc) => {
		const cached = procedures.get(proc);
		if (cached) { return cached; }
		let scope: SourceNameScope | undefined;
		const shapes = declarationShapeEnvironmentFor(symbols, proc);
		let types: ReturnType<typeof typeEnvironmentFor> | undefined;
		const callableShape = (name: string): DeclaredValueShape | undefined => {
			if (shapes.has(name)) { return undefined; }
			const definitions = declarations.get(name);
			const fn = definitions?.length === 1 ? definitions[0] : undefined;
			if (fn && (fn.kind === 'function' || fn.kind === 'propertyGet' || fn.kind === 'declare' && fn.declareKind === 'Function')) {
				return { asType: fn.asType, isArray: fn.isArray === true, isFixedArray: false };
			}
			if (definitions) { return undefined; }
			const runtime = resolveRuntimeFunction(name);
			if (runtime?.kind !== 'function' || (scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)) { return undefined; }
			return { asType: runtime.returns, isArray: false, isFixedArray: false };
		};
		const resolve = (expr: ExprNode): DeclaredValueShape | undefined => {
			if (expr.exprKind === 'ParenExpr') { return resolve(expr.inner); }
			if (expr.exprKind === 'IdentifierExpr') {
				const name = expr.name.toLowerCase();
				const shape = shapes.get(name);
				if (shape) { return shape; }
				// DefType also types undeclared assignment locals. They have no
				// declaration shape; reuse the effective type environment's binding
				// and host/runtime exclusions before attempting callable resolution.
				const asType = (types ??= typeEnvironmentFor(symbols, proc)).get(name);
				return asType ? { asType, isArray: false, isFixedArray: false } : callableShape(name);
			}
			if (expr.exprKind === 'MemberAccessExpr' && expr.object && expr.accessKind !== 'bang') {
				const holder = resolve(expr.object);
				if (!holder?.asType || holder.isArray) { return undefined; }
				const types = declarations.get(holder.asType.toLowerCase());
				const type = types?.length === 1 && types[0].kind === 'type' ? types[0] : undefined;
				let members = type ? fields.get(type) : undefined;
				if (type && !members) {
					const indexed = new Map<string, VbaSymbol>();
					for (const child of type.children ?? []) {
						const name = child.name.toLowerCase();
						if (!indexed.has(name)) { indexed.set(name, child); }
					}
					fields.set(type, indexed);
					members = indexed;
				}
				const field = members?.get(expr.member.toLowerCase());
				return field ? { asType: field.asType, isArray: field.isArray === true, isFixedArray: field.arrayBounds !== undefined } : undefined;
			}
			if (expr.exprKind === 'IndexExpr') {
				const calleeShape = expr.callee.exprKind === 'IdentifierExpr'
					? shapes.get(expr.callee.name.toLowerCase()) : resolve(expr.callee);
				if (calleeShape?.isArray) { return expr.args.length === 0 ? calleeShape : { ...calleeShape, isArray: false, isFixedArray: false }; }
				if (expr.callee.exprKind !== 'IdentifierExpr') { return undefined; }
				const result = callableShape(expr.callee.name.toLowerCase());
				// Arguments here invoke the function; they do not index its
				// returned array. Leave invalid invocations to argument-count.
				return result?.isArray && expr.args.length > 0 ? undefined : result;
			}
			return undefined;
		};
		procedures.set(proc, resolve);
		return resolve;
	};
	factories.set(symbols, { visible: projectVisibleSymbols, factory });
	return factory;
}
