// Rule family: duplicate and ambiguous declarations (audit #0).
//
// Extracted verbatim from analyzeModule.ts: duplicate procedures,
// declarations, module members, enum members, and ambiguous unqualified
// enum-member references.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { HostObjectModel } from '../../host/excelObjectModel';
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
	reportRepeatedKeys,
	valueReadReferences,
} from '../rules/shared';
import { extractCall } from '../callExtraction';
import {
	bareCallableSourceShadowed,
	callableTypeSignaturesFor,
	sameModuleCallableSignatures,
	sourceIdentifierBinding,
	sourceIdentifierBound,
	sourceNameScopeFor,
} from '../typeInference';
import { activeModuleMembers, type ProcedureStatementVisitor } from '../walker';

/** True when the symbol is one of the three Property accessor kinds. */
function isPropertyAccessor(sym: VbaSymbol): boolean {
	return sym.kind === 'propertyGet' || sym.kind === 'propertyLet' || sym.kind === 'propertySet';
}

/** What one of the name-collision rules below is looking for. */
interface RepeatedNameRule {
	/** Whether this symbol is one of the declarations the rule governs. */
	declares(sym: VbaSymbol): boolean;
	/** Whether an earlier declaration and a later one may not coexist. */
	collides(earlier: VbaSymbol, later: VbaSymbol): boolean;
	report(repeat: VbaSymbol): void;
}

/**
 * Reports every declaration that repeats a name an earlier one already took,
 * once per repeat, in source order.
 *
 * Two declarations in different arms of one `#If` chain never reach the
 * compiler together, so they are not a repeat however the conditional constants
 * evaluate. The arms XLIDE cannot decide are exactly where that matters: a
 * decidable condition has already left the losing arms inactive, and their
 * declarations never reach this rule.
 *
 * Almost every name is taken once, so a name holds its lone declaration
 * directly and grows a list only when a second one claims it. A name that DOES
 * collide stops at the first prior it collides with, so the cost is quadratic
 * only in how many same-name declarations exclude each other - that is, how
 * many arms one `#If` chain has. Real chains have a handful; it takes on the
 * order of a thousand arms of one chain, all declaring the same name, before
 * the scan is worth measuring.
 */
function reportRepeatedNames(
	symbols: readonly VbaSymbol[],
	activity: ConditionalActivityTracker | undefined,
	rule: RepeatedNameRule,
): void {
	const taken = new Map<string, VbaSymbol | VbaSymbol[]>();
	for (const sym of symbols) {
		if (!rule.declares(sym)) {
			continue;
		}
		const key = sym.name.toLowerCase();
		const holder = taken.get(key);
		if (holder === undefined) {
			taken.set(key, sym);
			continue;
		}
		const earlier = Array.isArray(holder) ? holder : [holder];
		const collision = earlier.some(
			(prior) =>
				rule.collides(prior, sym) &&
				!activity?.mutuallyExclusive(prior.nameSpan, sym.nameSpan),
		);
		if (collision) {
			rule.report(sym);
		}
		earlier.push(sym);
		taken.set(key, earlier);
	}
}

/**
 * Rule: a procedure name may name at most one Sub/Function, OR a set of distinct
 * Property accessors (one Get, one Let, one Set). Any other repeat is the VBA
 * "Ambiguous name detected" compile error.
 */
export function checkDuplicateProcedures(
	members: VbaSymbol[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	reportRepeatedNames(members, activity, {
		declares: (sym) => isProcedureKind(sym.kind),
		// Distinct accessors of one property share their name legitimately;
		// every other repeat is the ambiguity error.
		collides: (a, b) => !isPropertyAccessor(a) || !isPropertyAccessor(b) || a.kind === b.kind,
		report: (repeat) =>
			push(
				'duplicateProcedure',
				`Ambiguous name detected: '${repeat.name}' is already declared in this module.`,
				repeat.nameSpan,
			),
	});
}

/** Two declarations of one name always collide; only the scope differs. */
const DECLARATIONS_COLLIDE = (): boolean => true;

/**
 * Rule: within one procedure, a name may be declared once across its parameters,
 * local Dim/Static variables, and local Const declarations. Repeats are the VBA
 * "Duplicate declaration in current scope" error. Procedure scope is flat in VBA
 * (no block scope), so locals from different `If` branches still collide - but
 * locals from different `#If` branches do not, because only one of those arms
 * is ever compiled.
 */
export function checkDuplicateDeclarations(
	members: VbaSymbol[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const rule: RepeatedNameRule = {
		declares: (sym) =>
			sym.kind === 'parameter' || sym.kind === 'localVariable' || sym.kind === 'constant',
		collides: DECLARATIONS_COLLIDE,
		report: (repeat) =>
			push(
				'duplicateDeclaration',
				`Duplicate declaration in current scope: '${repeat.name}'.`,
				repeat.nameSpan,
			),
	};
	for (const proc of members) {
		if (isProcedureKind(proc.kind)) {
			reportRepeatedNames(proc.children ?? [], activity, rule);
		}
	}
}

/** Rule: a module-level variable or constant declared more than once. */
export function checkDuplicateModuleMembers(
	members: VbaSymbol[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	reportRepeatedNames(members, activity, {
		declares: (sym) => sym.kind === 'moduleVariable' || sym.kind === 'constant',
		collides: DECLARATIONS_COLLIDE,
		report: (repeat) =>
			push(
				'duplicateModuleMember',
				`Duplicate declaration: '${repeat.name}' is already declared at module level.`,
				repeat.nameSpan,
			),
	});
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
		// A member in a branch that cannot be decided still counts: it collides
		// with anything the same build would compile beside it, and only the
		// arms of one chain are alternatives. Skipping every undecidable branch
		// (which is what asking for `active` did) went blind to a genuine repeat
		// inside one arm.
		reportRepeatedKeys(member.members, activity, {
			keyOf: (entry) => (activity?.isInactive(entry.span) ? undefined : entry.name.toLowerCase()),
			spanOf: (entry) => entry.span,
			report: (repeat) => push(
				'duplicateEnumMember',
				`Duplicate Enum member '${repeat.name}' in Enum '${member.name}'.`,
				declarationNameHit(source, repeat.span, repeat.name)?.span ?? repeat.span,
			),
		});
	}
}

/** Rule: field names inside one Type (UDT) block must be unique. */
export function checkDuplicateTypeFields(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Type') {
			continue;
		}
		// See checkDuplicateEnumMembers: only provably-inactive fields drop out,
		// and the arms of one chain are alternatives rather than repeats.
		reportRepeatedKeys(member.fields, activity, {
			keyOf: (entry) => (activity?.isInactive(entry.span) ? undefined : entry.name.toLowerCase()),
			spanOf: (entry) => entry.span,
			report: (repeat) => push(
				'duplicateTypeField',
				`Duplicate field '${repeat.name}' in Type '${member.name}'.`,
				declarationNameHit(source, repeat.span, repeat.name)?.span ?? repeat.span,
			),
		});
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
	hostModel: HostObjectModel | undefined,
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
	// Presence test only: if no enum-member name is defined by more than one
	// visible Enum, no reference can be ambiguous, so skip the per-reference
	// binding work below. The groups map itself is not reused — each reference
	// re-derives ambiguity from its own resolved binding (which carries the
	// definitions), so caching the map here would not save that work.
	if (ambiguousEnumMemberGroups(visibleEnumMembers).size === 0) {
		return;
	}

	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const appMembers = applicationMemberNames(hostModel);
	const isKnownForSkip = (name: string, procSym: VbaSymbol | undefined): boolean => {
		const lower = name.toLowerCase();
		return (
			sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, 'expression') ||
			(knownProcedures?.has(lower) ?? false) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name, hostModel) !== undefined ||
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

/**
 * Rule: VBA is content for two modules to export the same public procedure
 * name, but it refuses to compile an UNQUALIFIED call to that name from a
 * module declaring neither - "Ambiguous name detected". Nothing reported it,
 * so a project the VBE will not compile read as clean.
 *
 * The finding belongs at the CALL SITE, not the declarations: a project that
 * exports a name twice and always qualifies its calls is legal VBA and common,
 * so flagging the declarations would cry wolf on every one of them.
 *
 * Silent, matching VBA, when any of these settle the name:
 *  - the call is qualified (`Helpers.Recalculate`)
 *  - the calling module declares the name itself (module-local scope wins)
 *  - a local, parameter or module-level symbol shadows it
 *  - only one module in the project exports it
 *
 * Per-statement rule: rides the shared procedure-statement walk (audit #0).
 */
export function checkAmbiguousBareProcedureCalls(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	moduleName: string,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const ambiguousNames = ambiguousProjectProcedureOwners(projectProcedures, moduleName);
	const sameModuleSignatures = sameModuleCallableSignatures(symbols);
	return (member) => {
		// No name in this project is exported twice: nothing here can be
		// ambiguous, so skip the per-statement work entirely.
		if (ambiguousNames.size === 0) {
			return () => { };
		}
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		return (stmt) => {
			const call = extractCall(source, stmt.span);
			if (!call || call.qualifier) {
				return;
			}
			const lower = call.name.toLowerCase();
			const owners = ambiguousNames.get(lower);
			if (!owners) {
				return;
			}
			// This module declares it, so VBA binds locally and never asks.
			if (sameModuleSignatures.has(lower)) {
				return;
			}
			if (bareCallableSourceShadowed(call.name, sourceNames)) {
				return;
			}
			push(
				'ambiguousProjectProcedure',
				`Ambiguous name detected: '${call.name}' is exported by ${owners.join(' and ')}. ` +
				'VBA refuses to compile the project until this call is qualified with a module name.',
				call.nameSpan,
			);
		};
	};
}

/**
 * Names exported by more than one OTHER module, mapped to those module names.
 * The calling module is excluded because its own declaration would settle the
 * name before the project is consulted.
 */
function ambiguousProjectProcedureOwners(
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	moduleName: string,
): Map<string, string[]> {
	const out = new Map<string, string[]>();
	if (!projectProcedures) {
		return out;
	}
	const self = moduleName.toLowerCase();
	for (const [name, signatures] of projectProcedures) {
		const owners: string[] = [];
		for (const signature of signatures) {
			// A Private procedure is not exported, so it cannot collide.
			if (signature.visibility === 'Private') {
				continue;
			}
			if (!owners.some((owner) => owner.toLowerCase() === signature.moduleName.toLowerCase())) {
				owners.push(signature.moduleName);
			}
		}
		if (owners.length > 1 && !owners.some((owner) => owner.toLowerCase() === self)) {
			out.set(name.toLowerCase(), owners);
		}
	}
	return out;
}
