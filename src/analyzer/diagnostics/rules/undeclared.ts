// Rule family: unresolved-name rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Option Explicit presence,
// undeclared variable reads/writes, member-not-found on exhaustively known
// surfaces, and unknown/non-callable bare call statements.

import {
	detectEol,
	VBA_IDENTIFIER_NAME_RE,
} from '../../../vbaSourceScan';
import { bareCallStatementTarget as callStatementTarget } from '../../call/callContext';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	resolveHostConstant,
	resolveHostGlobal,
} from '../../host/hostModel';
import { isReservedIdentifier } from '../../lexer/keywordTable';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	ModuleNode,
	Span,
} from '../../parser/nodes';
import {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
} from '../../runtime/vbaRuntime';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { BareIdentifierContext } from '../../symbols/nameResolution';
import type {
	VbaProcedureSignature,
	VbaProjectClassMembers,
	VbaSymbol,
} from '../../symbols/symbolModel';
import {
	applicationMemberNames,
	procedureSymbolFor,
	type PushFn,
	type VbaDiagnosticData,
} from '../analysisContext';
import {
	type CallableTypeSignature,
	type CallArguments,
	extractCall,
	isNamedSlot,
} from '../callExtraction';
import {
	forEachUndeclaredReferenceSpan,
	resolveExhaustiveMemberSurface,
	valueReadReferences,
} from '../rules/shared';
import {
	callableTypeSignaturesFor,
	isNonCallableSymbol,
	sourceIdentifierBinding,
	sourceIdentifierBound,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	matchParenFrom,
	setAssignmentTarget,
	statementTokens,
	tokenName,
	type ProcedureStatementVisitor,
} from '../walker';

/** Per-statement rule: rides the shared procedure-statement walk (audit #0). */
export function checkMemberNotFound(
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	return () => (stmt) => {
		for (const ref of memberAccessReferences(source, stmt.span)) {
			const surface = resolveExhaustiveMemberSurface(
				source,
				ref.dotEndOffset,
				memberCtx,
			);
			if (!surface) {
				continue;
			}
			if (
				surface.members.some(
					(candidate) =>
						candidate.name.toLowerCase() === ref.member.toLowerCase(),
				)
			) {
				continue;
			}
			push(
				'memberNotFound',
				`Method or data member not found: '${surface.owner}.${ref.member}'.`,
				ref.memberSpan,
			);
		}
	};
}

function memberAccessReferences(
	source: string,
	span: Span,
): { member: string; memberSpan: Span; dotEndOffset: number }[] {
	const toks = statementTokens(source, span);
	const out: { member: string; memberSpan: Span; dotEndOffset: number }[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (toks[i].rawText !== '.') {
			continue;
		}
		const member = tokenName(toks[i + 1]);
		if (!member) {
			continue;
		}
		out.push({
			member,
			memberSpan: {
				start: span.start + toks[i + 1].start,
				end: span.start + toks[i + 1].end,
			},
			dotEndOffset: span.start + toks[i].end,
		});
	}
	return out;
}

/**
 * Rule: a *call statement* whose callee is a bare (non-member) identifier - the
 * lone-identifier form `DoStartup`, the parenless-argument form `MsgBox "hi"` /
 * `Foo 1, 2`, or the explicit `Call DoWork` / `Call Foo(1, 2)` form - is a call
 * to a Sub/Function of that name. When the name resolves to nothing the VBE
 * raises "Sub or Function not defined".
 *
 * A name is considered resolved when it matches any project procedure, a name
 * declared in the current module (procedures, module variables/consts, types,
 * enums and their members, Declares), a parameter/local/const of the enclosing
 * procedure, a VBA runtime function/statement, or a host global / Application
 * member (Excel exposes Application's members in the global scope). The callee
 * detection ({@link callStatementTarget}) deliberately ignores assignments,
 * member calls, line labels, and the bare `Name(...)` indexed/implicit-member
 * form so those never produce a false positive.
 *
 * Per-statement rule: rides the shared procedure-statement walk (audit #0).
 */
export function checkUnknownCallStatement(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	knownProcedures: ReadonlySet<string>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	// Excel injects Application's members into the global scope, so a bare call
	// may legitimately bind to one of them (Calculate, Volatile, Evaluate, ...).
	const appMembers = applicationMemberNames();

	const isKnown = (name: string, procSym: VbaSymbol | undefined): boolean => {
		const lower = name.toLowerCase();
		return (
			knownProcedures.has(lower) ||
			sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, 'call') ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveRuntimeObject(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	return (member) => {
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			const hit = callStatementTarget(source, stmt.span);
			if (hit && !isKnown(hit.name, procSym)) {
				const call = extractCall(source, stmt.span);
				push(
					'unknownCallStatement',
					`Sub or Function not defined: '${hit.name}'.`,
					hit.span,
					call && call.nameSpan.start === hit.span.start && call.nameSpan.end === hit.span.end
						? createProcedureStubData(source, call)
						: undefined,
				);
			}
		};
	};
}

function createProcedureStubData(
	source: string,
	call: CallArguments,
): VbaDiagnosticData | undefined {
	if (!isGeneratedStubIdentifier(call.name)) {
		return undefined;
	}
	const params = generatedStubParameters(call);
	if (!params) {
		return undefined;
	}
	const eol = detectEol(source);
	const leading = source.length === 0
		? ''
		: `${source.endsWith('\n') || source.endsWith('\r') ? '' : eol}${endsWithBlankPhysicalLine(source) ? '' : eol}`;
	const text = `${leading}Private Sub ${call.name}(${params.join(', ')})${eol}End Sub${eol}`;
	return {
		createProcedureStub: {
			procedureName: call.name,
			edit: {
				span: { start: source.length, end: source.length },
				newText: text,
			},
		},
	};
}

function generatedStubParameters(call: CallArguments): string[] | undefined {
	if (call.slots.some((slot) => slot.length === 0)) {
		return undefined;
	}
	const named = call.slots.map((slot) => isNamedSlot(slot));
	if (named.some(Boolean) && !named.every(Boolean)) {
		return undefined;
	}
	const used = new Set<string>();
	const params: string[] = [];
	for (let i = 0; i < call.slots.length; i++) {
		const name = named[i]
			? generatedNamedArgumentParameterName(call.slots[i])
			: `arg${i + 1}`;
		if (!name || used.has(name.toLowerCase())) {
			return undefined;
		}
		used.add(name.toLowerCase());
		params.push(`ByVal ${name} As Variant`);
	}
	return params;
}

function generatedNamedArgumentParameterName(slot: VbaToken[]): string | undefined {
	const raw = slot[0]?.rawText;
	if (!raw || raw.startsWith('[')) {
		return undefined;
	}
	return isGeneratedStubIdentifier(raw) ? raw : undefined;
}

function isGeneratedStubIdentifier(name: string): boolean {
	return VBA_IDENTIFIER_NAME_RE.test(name) && !isReservedIdentifier(name);
}

function endsWithBlankPhysicalLine(source: string): boolean {
	return /(?:\r\n|\r|\n)[ \t]*(?:\r\n|\r|\n)$/.test(source);
}

/**
 * Rule: call statements must target a callable declaration. VBE Compile rejects
 * a bare non-callable statement (`testStr`), argument-bearing form
 * (`testStr "hello"`), and explicit `Call testStr` as call-shaped statements.
 */
export function checkNonCallableCallStatement(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	knownProcedures: ReadonlySet<string> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			const call = extractCall(source, stmt.span);
			if (!call) {
				return;
			}
			const binding = sourceIdentifierBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				call.name,
				'call',
			);
			if (binding.scope === 'ambiguous') {
				return;
			}
			if (binding.tier === 'project' && knownProcedures?.has(call.name.toLowerCase())) {
				return;
			}
			const target = binding.definitions.find((symbol) => isNonCallableSymbol(symbol));
			if (!target) {
				return;
			}
			if (callTargetFeedsMemberAccess(source, stmt.span, call)) {
				return;
			}
			push(
				'nonCallableCallStatement',
				`Cannot call '${call.name}' because it resolves to ${symbolKindLabel(target)}, not a Sub or Function.`,
				call.nameSpan,
			);
		};
	};
}

function callTargetFeedsMemberAccess(source: string, span: Span, call: CallArguments): boolean {
	const toks = statementTokens(source, span);
	const relCalleeStart = call.nameSpan.start - span.start;
	const calleeIdx = toks.findIndex((t) => t.start === relCalleeStart);
	if (calleeIdx < 0 || toks[calleeIdx + 1]?.rawText !== '(') {
		return false;
	}
	const close = matchParenFrom(toks, calleeIdx + 1);
	return close >= 0 && toks[close + 1]?.rawText === '.';
}

function symbolKindLabel(sym: VbaSymbol): string {
	switch (sym.kind) {
		case 'parameter':
			return 'a parameter';
		case 'localVariable':
			return 'a local variable';
		case 'moduleVariable':
			return 'a module variable';
		case 'constant':
			return 'a constant';
		case 'enum':
			return 'an enum type';
		case 'enumMember':
			return 'an enum member';
		case 'type':
			return 'a user-defined type';
		default:
			return 'a non-callable declaration';
	}
}

/**
 * Rule: a code module that contains real code but no `Option Explicit` lets
 * variables be used without declaration. Empty/attribute-only modules are
 * skipped to avoid noise on blank document modules.
 */
export function checkOptionExplicit(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let hasExplicit = false;
	let hasCode = false;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Option' && /^explicit\b/i.test(member.optionText.trim())) {
			hasExplicit = true;
		}
		if (
			member.kind === 'Procedure' ||
			member.kind === 'VariableGroup' ||
			member.kind === 'Type' ||
			member.kind === 'Enum' ||
			member.kind === 'Declare'
		) {
			hasCode = true;
		}
	}
	if (hasExplicit || !hasCode) {
		return;
	}
	// Anchor as a zero-width marker at the module top - the insertion point for
	// `Option Explicit` - rather than spanning the whole first physical line.
	// A full first-line range collides with any error on that line (e.g. the
	// missing-block-closer on a not-yet-closed first procedure, a common editing
	// state), and a warning sharing an error's exact range obscures the red
	// squiggle. A zero-width range keeps the gutter/Problems entry without
	// painting over a more severe diagnostic.
	push(
		'optionExplicitMissing',
		'Option Explicit is not specified; variables can be used without being declared. Add "Option Explicit" to the top of the module.',
		{ start: 0, end: 0 },
	);
}

/**
 * Rule: with `Option Explicit`, a variable must be declared before it can be
 * assigned or read. The rule only runs once the caller has supplied the
 * project-visible identifier set, so cross-module globals and enum members do
 * not false-positive.
 */
export function checkUndeclaredVariables(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	knownIdentifiers: ReadonlySet<string> | undefined,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): void {
	if (!hasOptionExplicit(mod, activity) || !knownIdentifiers) {
		return;
	}

	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const appMembers = applicationMemberNames();
	const isKnown = (
		name: string,
		procSym: VbaSymbol | undefined,
		context: BareIdentifierContext,
	): boolean => {
		const lower = name.toLowerCase();
		return (
			lower === 'vba' ||
			sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, context) ||
			knownIdentifiers.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveHostConstant(name) !== undefined ||
			resolveRuntimeConstant(name) !== undefined ||
			resolveRuntimeObject(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = procedureSymbolFor(symbols, member);
		forEachUndeclaredReferenceSpan(source, member.body, (span) => {
			const reported = new Set<string>();
			const report = (
				name: string,
				span: Span,
				mode: 'assigning to it' | 'using it',
				context: BareIdentifierContext,
			): void => {
				const key = `${span.start}:${span.end}`;
				if (reported.has(key) || isKnown(name, procSym, context)) {
					return;
				}
				reported.add(key);
				push(
					'undeclaredVariable',
					`Variable not defined: '${name}'. Declare it before ${mode}, or remove Option Explicit.`,
					span,
				);
			};
			const scalarTarget = bareAssignmentTarget(source, span);
			const objectTarget = scalarTarget ? undefined : setAssignmentTarget(source, span);
			const target = scalarTarget ?? objectTarget;
			if (target) {
				report(target.name, target.span, 'assigning to it', 'assignmentTarget');
			}
			for (const ref of undeclaredReadReferences(
				source,
				span,
				(name) => isKnown(name, procSym, 'expression'),
				moduleSignatures,
				projectMembers,
			)) {
				report(ref.name, ref.span, 'using it', 'expression');
			}
		}, activity);
	}
}

function undeclaredReadReferences(
	source: string,
	span: Span,
	isKnown: (name: string) => boolean,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): Array<{ name: string; span: Span }> {
	return valueReadReferences(source, span, isKnown, moduleSignatures, projectMembers)
		.filter((ref) => !isKnown(ref.name));
}

function hasOptionExplicit(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): boolean {
	return activeModuleMembers(mod, activity).some(
		(member) =>
			member.kind === 'Option' && /^explicit\b/i.test(member.optionText.trim()),
	);
}
