// Conservative declaration shapes for operands. Resolve only declarations we
// own; Object, Variant, unresolved members and ambiguous call results stay unknown.
import type { Argument, ExprNode, ProcedureNode } from '../parser/nodes';
import { qualifiedProcedureKey, type VbaSymbol, type VbaProcedureSignature } from '../symbols/symbolModel';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { resolveRuntimeConstant, resolveRuntimeFunction, resolveRuntimeObject, type VbaRuntimeFunction } from '../runtime/vbaRuntime';
import { isNumericType, moduleSignatureFromSymbol, uniqueProjectTypeSignatures, runtimeAritySignature, declarationShapeEnvironmentFor, defTypeOf, typeEnvironmentFor, sourceNameScopeFor, type SourceNameScope, type DeclaredValueShape } from './typeInference';
import { VBA_LIBRARY_CONTAINERS } from '../runtime/vbaLibraryNames';
import type { CallableTypeSignature } from './callExtraction';
import type { MemberCompletionContext } from '../completion/memberAccess';
import { resolveHostConstant, resolveHostGlobal, resolveHostGlobalMember } from '../host/hostModel';

export type ExpressionShapeContext = Pick<MemberCompletionContext, 'projectClassMembers' | 'implicitMembers'> & {
	/** Caller-supplied exported signatures, with declaring-module effective returns. */
	projectProcedures?: ReadonlyMap<string, readonly VbaProcedureSignature[]>;
};

type ShapeFactory = (proc: ProcedureNode) => (expr: ExprNode) => DeclaredValueShape | undefined;
const factories = new WeakMap<ReturnType<typeof buildModuleSymbols>, { visible: readonly VbaSymbol[] | undefined; context: ExpressionShapeContext | undefined; factory: ShapeFactory }>();

export function expressionShapeResolver(symbols: ReturnType<typeof buildModuleSymbols>, projectVisibleSymbols?: readonly VbaSymbol[], context?: ExpressionShapeContext): ShapeFactory {
	const cached = factories.get(symbols);
	if (cached && cached.visible === projectVisibleSymbols && cached.context === context) { return cached.factory; }
	// Indexed once per analysis, never scanned for each operand. These names
	// have project/designer bindings and cannot become DefType implicit locals.
	let nonLocalNames: Set<string> | undefined;
	const valueBindings = new Set<string>();
	const isNonLocalName = (name: string): boolean => {
		if (!nonLocalNames) {
			nonLocalNames = new Set(['me', symbols.moduleName.toLowerCase()]);
			for (const type of context?.projectClassMembers ?? []) {
				const name = type.name.toLowerCase();
				nonLocalNames.add(name);
				// An unread attribute header cannot prove the absence of a default instance.
				if (type.kind === 'document' || type.kind === 'userform' || type.kind === 'class' && type.predeclaredId !== false) {
					valueBindings.add(name);
				}
			}
			for (const member of context?.implicitMembers ?? []) {
				const name = member.name.toLowerCase();
				nonLocalNames.add(name);
				valueBindings.add(name);
			}
			for (const key of context?.projectProcedures?.keys() ?? []) { if (key.includes('.')) { nonLocalNames.add(key.split('.')[0]); } }
			for (const symbol of projectVisibleSymbols ?? []) { nonLocalNames.add(symbol.moduleName.toLowerCase()); }
		}
		return nonLocalNames.has(name);
	};
	const declarations = new Map<string, NonNullable<typeof symbols.root.children>>();
	for (const symbol of symbols.root.children ?? []) {
		const key = symbol.name.toLowerCase();
		const group = declarations.get(key) ?? [];
		group.push(symbol);
		declarations.set(key, group);
	}
	const readDeclarations = new Map<string, VbaSymbol>();
	for (const [name, group] of declarations) {
		// A read selects the Get of a property family; its Let/Set are writers.
		const readers = group.filter((symbol) => symbol.kind !== 'propertyLet' && symbol.kind !== 'propertySet');
		if (readers.length === 1 && (group.length === 1 || readers[0].kind === 'propertyGet')) { readDeclarations.set(name, readers[0]); }
	}
	const projectSignatures = uniqueProjectTypeSignatures(context?.projectProcedures);
	const projectShape = (key: string, args: readonly Argument[]): DeclaredValueShape | undefined => {
		const signature = projectSignatures.get(key);
		if (!signature?.valued || !acceptsArguments(signature, args)) { return undefined; }
		const isArray = /\(\s*\)\s*$/.test(signature.returnType ?? '');
		return { asType: signature.returnType, isArray, isFixedArray: false };
	};
	const signatures = new WeakMap<VbaSymbol, CallableTypeSignature>();
	const runtimeSignatures = new WeakMap<VbaRuntimeFunction, CallableTypeSignature | undefined>();
	const runtimeShape = (runtime: VbaRuntimeFunction, args: readonly Argument[], scalarArgument: (expr: ExprNode) => boolean): DeclaredValueShape | undefined => {
		if (!runtimeSignatures.has(runtime)) { runtimeSignatures.set(runtime, runtimeAritySignature(runtime)); }
		const signature = runtimeSignatures.get(runtime);
		if (signature && !acceptsArguments(signature, args)) { return undefined; }
		// Abs/Int/Fix constrain numeric/String arguments, while Null, Variant and
		// Object inputs compiled in the 2026-10-19 Excel probes. This is an
		// invocation constraint, not proof of an exact return declaration.
		const conditionalScalar = /^(abs|int|fix)$/i.test(runtime.name);
		const argument = args[0]?.value;
		const isScalar = runtime.scalarReturn && (!conditionalScalar || !!argument && scalarArgument(argument));
		return { asType: runtime.declaredReturns, isScalar, isArray: false, isFixedArray: false };
	};
	const fields = new WeakMap<VbaSymbol, ReadonlyMap<string, VbaSymbol>>();
	const procedures = new WeakMap<ProcedureNode, ReturnType<ShapeFactory>>();
	const factory: ShapeFactory = (proc) => {
		const cached = procedures.get(proc);
		if (cached) { return cached; }
		let scope: SourceNameScope | undefined;
		const shapes = declarationShapeEnvironmentFor(symbols, proc);
		let types: ReturnType<typeof typeEnvironmentFor> | undefined;
		const scalarArgument = (expr: ExprNode): boolean => {
			if (expr.exprKind === 'ParenExpr') { return scalarArgument(expr.inner); }
			if (expr.exprKind === 'UnaryExpr' && (expr.operator === '+' || expr.operator === '-')) { return scalarArgument(expr.operand); }
			if (expr.exprKind === 'LiteralExpr') { return expr.literalKind === 'integer' || expr.literalKind === 'float' || expr.literalKind === 'string'; }
			const shape = resolve(expr);
			return !shape?.isArray && !!shape?.asType && (isNumericType(shape.asType.toLowerCase()) || shape.asType.toLowerCase() === 'string');
		};
		const callableShape = (name: string, args: readonly Argument[] = [], suffix = '', moduleQualified = false): DeclaredValueShape | undefined => {
			if (!moduleQualified && shapes.has(name)) { return undefined; }
			const definitions = declarations.get(name);
			const fn = readDeclarations.get(name);
			if (fn && (fn.kind === 'function' || fn.kind === 'propertyGet' || fn.kind === 'declare' && fn.declareKind === 'Function')) {
				let signature = signatures.get(fn);
				if (!signature) { signatures.set(fn, signature = moduleSignatureFromSymbol(symbols, fn)); }
				if (!acceptsArguments(signature, args)) { return undefined; }
				return { asType: signature.returnType, isArray: fn.isArray === true, isFixedArray: false };
			}
			if (!definitions && !moduleQualified && context?.projectProcedures?.has(name)) {
				return isNonLocalName(name) && valueBindings.has(name) ? undefined : projectShape(name, args);
			}
			if (definitions || moduleQualified || isNonLocalName(name) && valueBindings.has(name)) { return undefined; }
			const runtime = resolveRuntimeFunction(name + suffix);
			if (runtime?.kind !== 'function' || (scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)) { return undefined; }
			return runtimeShape(runtime, args, scalarArgument);
		};
		const callableQualifier = (expr: ExprNode): 'module' | 'vba' | 'project' | undefined => {
			if (expr.exprKind !== 'MemberAccessExpr' || expr.accessKind === 'bang') { return undefined; }
			const object = expr.object;
			if (object?.exprKind === 'MemberAccessExpr' && object.accessKind !== 'bang'
				&& object.object?.exprKind === 'IdentifierExpr' && object.object.name.toLowerCase() === 'vba'
				&& symbols.moduleName.toLowerCase() !== 'vba'
				&& !shapes.has('vba') && !declarations.has('vba') && !(isNonLocalName('vba') && valueBindings.has('vba'))
				&& !(scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has('vba')
				&& VBA_LIBRARY_CONTAINERS.get(object.member.toLowerCase())?.has(expr.member.toLowerCase())) { return 'vba'; }
			if (object?.exprKind !== 'IdentifierExpr') { return undefined; }
			const name = object.name.toLowerCase();
			if (shapes.has(name) || declarations.has(name) || isNonLocalName(name) && valueBindings.has(name)) { return undefined; }
			if (symbols.moduleKind === 'standard' && name === symbols.moduleName.toLowerCase()) { return 'module'; }
			if (context?.projectProcedures?.has(qualifiedProcedureKey(name, expr.member.replace(/[$%&!#@^]$/, '')))
				&& !(scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)) { return 'project'; }
			if (!isNonLocalName(name) && VBA_LIBRARY_CONTAINERS.get(name)?.has(expr.member.toLowerCase())
				&& !(scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)) { return 'vba'; }
			if (name === 'vba' && !(scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)) { return 'vba'; }
			return undefined;
		};
		const qualifiedCallableShape = (expr: Extract<ExprNode, { exprKind: 'MemberAccessExpr' }>, qualifier: 'module' | 'vba' | 'project', args: readonly Argument[] = []): DeclaredValueShape | undefined => {
			// Member syntax retains its suffix and span; declaration keys omit it.
			const name = expr.member.replace(/[$%&!#@^]$/, '').toLowerCase();
			if (qualifier === 'module') { return callableShape(name, args, '', true); }
			if (qualifier === 'project' && expr.object?.exprKind === 'IdentifierExpr') { return projectShape(qualifiedProcedureKey(expr.object.name, name), args); }
			const runtime = resolveRuntimeFunction(expr.member);
			if (runtime?.kind !== 'function') { return undefined; }
			return runtimeShape(runtime, args, scalarArgument);
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
				const asType = !isNonLocalName(name) && !declarations.has(name)
					? (types ??= typeEnvironmentFor(symbols, proc)).get(name) : undefined;
				if (asType) { return { asType, isArray: false, isFixedArray: false }; }
				const callable = callableShape(name, [], expr.typeSuffix);
				if (callable) { return callable; }
				// A read alone also creates an implicit local without Option Explicit.
				// Preserve source/project, host and runtime bindings before DefType.
				const implicitType = symbols.implicitLocals !== undefined && !expr.typeSuffix ? defTypeOf(symbols, name) : undefined;
				return implicitType && !isNonLocalName(name) && !(scope ??= sourceNameScopeFor(symbols, proc, projectVisibleSymbols)).runtimeShadows.has(name)
					&& !resolveRuntimeFunction(name) && !resolveRuntimeObject(name) && !resolveRuntimeConstant(name)
					&& !resolveHostGlobal(name, undefined) && !resolveHostGlobalMember(name, undefined) && !resolveHostConstant(name, undefined)
					? { asType: implicitType, isArray: false, isFixedArray: false } : undefined;
			}
			if (expr.exprKind === 'MemberAccessExpr' && expr.object && expr.accessKind !== 'bang') {
				const qualifier = callableQualifier(expr);
				if (qualifier) { return qualifiedCallableShape(expr, qualifier); }
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
				const qualifier = callableQualifier(expr.callee);
				if (qualifier && expr.callee.exprKind === 'MemberAccessExpr') { return qualifiedCallableShape(expr.callee, qualifier, expr.args); }
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
	factories.set(symbols, { visible: projectVisibleSymbols, context, factory });
	return factory;
}

/** Match argument slots to the declaration before judging its result shape. */
function acceptsArguments(signature: CallableTypeSignature, args: readonly Argument[]): boolean {
	// InStr's two-positional-argument form omits the leading Start slot.
	const params = signature.runtime && /^(instr|instrb)$/i.test(signature.name) && args.length === 2 && args.every((arg) => !arg.name)
		? signature.params.slice(1) : signature.params;
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
