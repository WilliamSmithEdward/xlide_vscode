// Rule family: duplicate and ambiguous declarations (audit #0).
//
// Extracted verbatim from analyzeModule.ts: duplicate procedures,
// declarations, module members, enum members, and ambiguous unqualified
// enum-member references.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { resolveHostGlobal } from '../../host/hostModel';
import type { ModuleNode } from '../../parser/nodes';
import {
	resolveRuntimeFunction,
	resolveRuntimeObject,
} from '../../runtime/vbaRuntime';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { BareIdentifierResolution } from '../../symbols/nameResolution';
import {
	isProcedureKind,
	type VbaProcedureSignature,
	type VbaProjectClassMembers,
	type VbaSymbol,
} from '../../symbols/symbolModel';
import {
	applicationMemberNames,
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import {
	declarationNameHit,
	forEachUndeclaredReferenceSpan,
	valueReadReferences,
} from '../rules/shared';
import {
	callableTypeSignaturesFor,
	sourceIdentifierBinding,
	sourceIdentifierBound,
} from '../typeInference';
import { activeModuleMembers } from '../walker';

/**
 * Rule: a procedure name may name at most one Sub/Function, OR a set of distinct
 * Property accessors (one Get, one Let, one Set). Any other repeat is the VBA
 * "Ambiguous name detected" compile error.
 */
export function checkDuplicateProcedures(members: VbaSymbol[], push: PushFn): void {
	const groups = new Map<string, VbaSymbol[]>();
	for (const sym of members) {
		if (!isProcedureKind(sym.kind)) {
			continue;
		}
		const key = sym.name.toLowerCase();
		(groups.get(key) ?? groups.set(key, []).get(key)!).push(sym);
	}

	for (const group of groups.values()) {
		if (group.length < 2) {
			continue;
		}
		let valueProcSeen = false;
		const accessorSeen = new Set<string>();
		for (const sym of group) {
			const isProperty =
				sym.kind === 'propertyGet' ||
				sym.kind === 'propertyLet' ||
				sym.kind === 'propertySet';
			let conflict = false;
			if (!isProperty) {
				conflict = valueProcSeen || accessorSeen.size > 0;
				valueProcSeen = true;
			} else {
				conflict = valueProcSeen || accessorSeen.has(sym.kind);
				accessorSeen.add(sym.kind);
			}
			if (conflict) {
				push(
					'duplicateProcedure',
					`Ambiguous name detected: '${sym.name}' is already declared in this module.`,
					sym.nameSpan,
				);
			}
		}
	}
}

/**
 * Rule: within one procedure, a name may be declared once across its parameters,
 * local Dim/Static variables, and local Const declarations. Repeats are the VBA
 * "Duplicate declaration in current scope" error. Procedure scope is flat in VBA
 * (no block scope), so locals from different branches still collide.
 */
export function checkDuplicateDeclarations(members: VbaSymbol[], push: PushFn): void {
	for (const proc of members) {
		if (!isProcedureKind(proc.kind)) {
			continue;
		}
		const seen = new Set<string>();
		for (const child of proc.children ?? []) {
			if (
				child.kind !== 'parameter' &&
				child.kind !== 'localVariable' &&
				child.kind !== 'constant'
			) {
				continue;
			}
			const key = child.name.toLowerCase();
			if (seen.has(key)) {
				push(
					'duplicateDeclaration',
					`Duplicate declaration in current scope: '${child.name}'.`,
					child.nameSpan,
				);
			} else {
				seen.add(key);
			}
		}
	}
}

/** Rule: a module-level variable or constant declared more than once. */
export function checkDuplicateModuleMembers(members: VbaSymbol[], push: PushFn): void {
	const seen = new Set<string>();
	for (const sym of members) {
		if (sym.kind !== 'moduleVariable' && sym.kind !== 'constant') {
			continue;
		}
		const key = sym.name.toLowerCase();
		if (seen.has(key)) {
			push(
				'duplicateModuleMember',
				`Duplicate declaration: '${sym.name}' is already declared at module level.`,
				sym.nameSpan,
			);
		} else {
			seen.add(key);
		}
	}
}

/** Rule: member names inside one Enum block must be unique. */
export function checkDuplicateEnumMembers(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Enum') {
			continue;
		}
		const seen = new Set<string>();
		for (const enumMember of member.members) {
			// Only provably-active members can collide. A member in an inactive
			// branch is never compiled, and a member in a not-provably-active
			// branch (an `#If` on an unknown constant - including the two arms of
			// an `#If`/`#Else`) is not guaranteed to be compiled alongside another
			// same-named member, so reporting it as a duplicate would be a false
			// positive. When there are no directives, `activity` is undefined and
			// every member counts as active.
			if (activity && activity.activityForSpan(enumMember.span) !== 'active') {
				continue;
			}
			const key = enumMember.name.toLowerCase();
			const hit = declarationNameHit(source, enumMember.span, enumMember.name);
			if (seen.has(key)) {
				push(
					'duplicateEnumMember',
					`Duplicate Enum member '${enumMember.name}' in Enum '${member.name}'.`,
					hit?.span ?? enumMember.span,
				);
			} else {
				seen.add(key);
			}
		}
	}
}

/**
 * Rule: duplicate member names in different Enum blocks compile, but an
 * unqualified read of that shared member name is rejected as "Ambiguous name
 * detected". Same-module bindings take precedence over exported members from
 * other modules, and procedure locals/parameters shadow module-level enum
 * members.
 */
export function checkAmbiguousEnumMemberReferences(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	moduleName: string,
	knownProcedures: ReadonlySet<string> | undefined,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): void {
	const visibleEnumMembers = [
		...enumMemberSymbols(symbols.root.children ?? []),
		...(projectVisibleSymbols ?? []).filter(
			(sym) =>
				sym.kind === 'enumMember' &&
				sym.moduleName.toLowerCase() !== moduleName.toLowerCase(),
		),
	];
	if (ambiguousEnumMemberGroups(visibleEnumMembers).size === 0) {
		return;
	}

	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const appMembers = applicationMemberNames();
	const isKnownForSkip = (name: string, procSym: VbaSymbol | undefined): boolean => {
		const lower = name.toLowerCase();
		return (
			sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, 'expression') ||
			(knownProcedures?.has(lower) ?? false) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveRuntimeObject(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = procedureSymbolFor(symbols, member);
		const reported = new Set<string>();
		forEachUndeclaredReferenceSpan(source, member.body, (span) => {
			for (const ref of valueReadReferences(
				source,
				span,
				(name) => isKnownForSkip(name, procSym),
				moduleSignatures,
				projectMembers,
			)) {
				const binding = sourceIdentifierBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					ref.name,
					'expression',
				);
				const definitions = ambiguousEnumMemberDefinitions(binding);
				if (!definitions) {
					continue;
				}
				const key = `${ref.span.start}:${ref.span.end}`;
				if (reported.has(key)) {
					continue;
				}
				reported.add(key);
				const owners = definitions
					.map((definition) => definition.containerName ?? definition.moduleName)
					.filter((owner, index, all) => all.indexOf(owner) === index)
					.slice(0, 3)
					.join(', ');
				push(
					'ambiguousEnumMember',
					`Ambiguous Enum member reference: '${ref.name}' is defined by multiple visible Enums${owners ? ` (${owners})` : ''}. Qualify the reference with an Enum or module name.`,
					ref.span,
				);
			}
		}, activity);
	}
}

function enumMemberSymbols(symbols: readonly VbaSymbol[]): VbaSymbol[] {
	const out: VbaSymbol[] = [];
	for (const symbol of symbols) {
		if (symbol.kind === 'enum') {
			out.push(...(symbol.children ?? []).filter((child) => child.kind === 'enumMember'));
		}
	}
	return out;
}

function ambiguousEnumMemberGroups(symbols: readonly VbaSymbol[]): Map<string, VbaSymbol[]> {
	const groups = new Map<string, Map<string, VbaSymbol>>();
	for (const symbol of symbols) {
		if (symbol.kind !== 'enumMember') {
			continue;
		}
		const key = symbol.name.toLowerCase();
		const ownerKey = `${symbol.moduleName.toLowerCase()}:${(symbol.containerName ?? '').toLowerCase()}`;
		const owners = groups.get(key) ?? new Map<string, VbaSymbol>();
		if (!owners.has(ownerKey)) {
			owners.set(ownerKey, symbol);
		}
		groups.set(key, owners);
	}
	const ambiguous = new Map<string, VbaSymbol[]>();
	for (const [key, owners] of groups) {
		if (owners.size > 1) {
			ambiguous.set(key, [...owners.values()]);
		}
	}
	return ambiguous;
}

function ambiguousEnumMemberDefinitions(
	binding: BareIdentifierResolution,
): readonly VbaSymbol[] | undefined {
	if (binding.scope !== 'ambiguous') {
		return undefined;
	}
	if (binding.definitions.some((definition) => definition.kind !== 'enumMember')) {
		return undefined;
	}
	const ownerKeys = new Set(
		binding.definitions.map(
			(definition) =>
				`${definition.moduleName.toLowerCase()}:${(definition.containerName ?? '').toLowerCase()}`,
		),
	);
	return ownerKeys.size > 1 ? binding.definitions : undefined;
}
