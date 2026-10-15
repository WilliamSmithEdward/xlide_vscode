// Conservative declaration shapes for operands. Resolve only declarations we
// own; Object, Variant, unresolved members and ambiguous call results stay unknown.
import type { Argument, ExprNode, ProcedureNode } from '../parser/nodes';
import type { VbaSymbol } from '../symbols/symbolModel';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import { callableTypeSignatureFromSymbol, declarationShapeEnvironmentFor, typeEnvironmentFor, sourceNameScopeFor, type SourceNameScope, type DeclaredValueShape } from './typeInference';
import type { CallableTypeSignature } from './callExtraction';

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
	const signatures = new WeakMap<VbaSymbol, CallableTypeSignature>();
	const fields = new WeakMap<VbaSymbol, ReadonlyMap<string, VbaSymbol>>();
	const procedures = new WeakMap<ProcedureNode, ReturnType<ShapeFactory>>();
	const factory: ShapeFactory = (proc) => {
		const cached = procedures.get(proc);
		if (cached) { return cached; }
		let scope: SourceNameScope | undefined;
		const shapes = declarationShapeEnvironmentFor(symbols, proc);
		let types: ReturnType<typeof typeEnvironmentFor> | undefined;
		const callableShape = (name: string, args: readonly Argument[] = [], suffix = ''): DeclaredValueShape | undefined => {
			if (shapes.has(name)) { return undefined; }
			const definitions = declarations.get(name);
			const fn = definitions?.length === 1 ? definitions[0] : undefined;
			if (fn && (fn.kind === 'function' || fn.kind === 'propertyGet' || fn.kind === 'declare' && fn.declareKind === 'Function')) {
				if (fn.isArray) {
					let signature = signatures.get(fn);
					if (!signature) { signatures.set(fn, signature = callableTypeSignatureFromSymbol(fn)); }
					if (!acceptsArguments(signature, args)) { return undefined; }
				}
				return { asType: fn.asType, isArray: fn.isArray === true, isFixedArray: false };
			}
			if (definitions) { return undefined; }
			const runtime = resolveRuntimeFunction(name + suffix);
			if (runtime?.kind !== 'function' || (scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)) { return undefined; }
			return { asType: runtime.declaredReturns, isArray: false, isFixedArray: false };
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
				return asType ? { asType, isArray: false, isFixedArray: false } : callableShape(name, [], expr.typeSuffix);
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
				// Arguments invoke the function, rather than indexing its result.
				// Only invalid invocations belong exclusively to argument-count.
				return callableShape(expr.callee.name.toLowerCase(), expr.args, expr.callee.typeSuffix);
			}
			return undefined;
		};
		procedures.set(proc, resolve);
		return resolve;
	};
	factories.set(symbols, { visible: projectVisibleSymbols, factory });
	return factory;
}

/** Match argument slots to the declaration before judging an array result. */
function acceptsArguments(signature: CallableTypeSignature, args: readonly Argument[]): boolean {
	const params = signature.params;
	const supplied = new Set<number>();
	let named = false;
	for (let i = 0; i < args.length; i++) {
		const arg = args[i];
		let index = i;
		if (arg.name) {
			named = true;
			if (params.some((param) => param.paramArray)) { return false; }
			index = params.findIndex((param) => param.name.toLowerCase() === arg.name!.toLowerCase());
		} else if (named) { return false; }
		const param = params[index] ?? (params.at(-1)?.paramArray ? params.at(-1) : undefined);
		if (!param || supplied.has(index)) { return false; }
		if (!arg.value) {
			if ((!param.optional && !param.paramArray) || i === args.length - 1) { return false; }
		} else { supplied.add(index); }
	}
	return params.every((param, index) => param.optional || param.paramArray || supplied.has(index));
}
