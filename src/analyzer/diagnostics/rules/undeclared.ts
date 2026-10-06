// Rule family: unresolved-name rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Option Explicit presence,
// undeclared variable reads/writes, member-not-found on exhaustively known
// surfaces, and unknown/non-callable bare call statements.

import {
	detectEol,
	lineStartAtAnyBreak,
	lineEndAtOrAfter,
	VBA_IDENTIFIER_NAME_RE,
} from '../../../vbaSourceScan';
import { getExcelObjectModel, type HostObjectModel } from '../../host/excelObjectModel';
import { HOST_LIBRARY_NAMES } from '../../host/hostLibraries';
import type { VbaHostToken } from '../../host/hostRegistry';
import { bareCallStatementTarget as callStatementTarget } from '../../call/callContext';
import { privateMemberOwnerAt, projectClassMemberAt, projectTypeAt, resolveMemberPresenceSurfaceAt, type MemberCompletionContext } from '../../completion/memberAccess';
import { MSFORMS_FORM_CONTROL_MEMBERS } from '../../host/msFormsFormControlMembers';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	resolveHostConstant,
	resolveHostEnum,
	resolveHostGlobal,
	resolveHostGlobalMember,
} from '../../host/hostModel';
import { isReservedIdentifier } from '../../lexer/keywordTable';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ModuleNode,
	ProcedureNode,
	Span,
} from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import {
	resolveExpressionType,
	type ExpressionTypeContext,
} from '../../expression/resolveExpressionType';
import {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
	resolveVbaLibraryQualifier,
	type VbaRuntimeFunction,
} from '../../runtime/vbaRuntime';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { BareIdentifierContext } from '../../symbols/nameResolution';
import {
	type ModuleSymbolKind,
	type VbaProcedureSignature,
	type VbaProjectClassMembers,
	type VbaSymbol,
} from '../../symbols/symbolModel';
import {
	applicationMemberNames,
	designerClassMemberNames,
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
	valueReadReferences,
} from '../rules/shared';
import {
	callableTypeSignaturesFor,
	isKnownScalarType,
	isNonCallableSymbol,
	normalizeType,
	sourceIdentifierBinding,
	sourceIdentifierBound,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	forEachStatement,
	forEachVariableGroup,
	matchParenFrom,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/**
 * Rule: a VBA library procedure named bare where a value goes (issue #318,
 * measured in Excel 16.0, with or without Option Explicit). One that needs an
 * argument, `Main = Left` or `TypeName(Kill)`, is "Argument not optional"; a
 * statement that takes none, `Main = Beep` or `Reset`, is "Expected Function
 * or variable". One whose arguments are all optional, Now or Timer, gives its
 * value. A name the module, the project, the host or the module's own object
 * declares is theirs.
 */
export function checkBuiltinsReadBare(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	moduleKind: ModuleSymbolKind | undefined,
	hostModel: HostObjectModel | undefined,
	designerClass: string | undefined,
	implicitMembers: readonly { name: string; type: string }[] | undefined,
	push: PushFn,
	ownMembers: ReadonlySet<string> = new Set(),
): void {
	if (moduleKind === 'userform' && implicitMembers === undefined) {
		return;
	}
	const appMembers = applicationMemberNames(hostModel);
	const designerMembers = designerClassMemberNames(designerClass, hostModel);
	const implicitNames = new Set((implicitMembers ?? []).map((member) => member.name.toLowerCase()));
	const explicit = hasOptionExplicit(mod, activity);
	const builtin = (name: string, procSym: VbaSymbol | undefined): VbaRuntimeFunction | undefined => {
		const lower = name.toLowerCase();
		if (appMembers.has(lower) || designerMembers.has(lower) || ownMembers.has(lower) || implicitNames.has(lower)
			|| sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, 'expression')
			|| resolveHostGlobal(name, hostModel) !== undefined || resolveHostGlobalMember(name, hostModel) !== undefined
			|| resolveHostConstant(name, hostModel) !== undefined || resolveHostEnum(name, hostModel) !== undefined
			|| resolveRuntimeConstant(name) !== undefined || resolveRuntimeObject(name) !== undefined) {
			return undefined;
		}
		return resolveRuntimeFunction(name);
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = procedureSymbolFor(symbols, member);
		const redim = redimTargetNamesIn(source, member.body, activity);
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const assignment = bareAssignmentTarget(source, span);
				const valueFrom = assignment && assignment.valueTokens.length > 0 ? toks.findIndex((tok) => tok.start === assignment.valueTokens[0].start) : -1;
				let depth = 0;
				toks.forEach((tok, i) => {
					depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
					const next = toks[i + 1]?.rawText;
					// A name after AddressOf is addressof-misuse's (issue #299).
					if (tok.kind !== 'identifier' || (depth === 0 && (valueFrom < 0 || i < valueFrom)) || ['.', '!'].includes(toks[i - 1]?.rawText ?? '') || tokenText(toks[i - 1]) === 'addressof'
						|| ['(', '.', '!', ':=', '$'].includes(next ?? '') || redim.has(tok.rawText.toLowerCase())) {
						return;
					}
					const runtime = builtin(tok.rawText, procSym);
					if (!runtime) {
						return;
					}
					const at = { start: span.start + tok.start, end: span.start + tok.end };
					// `Line` and `Name` open statements and name no procedure: read as a
					// value under Option Explicit, each is "Variable not defined".
					if (runtime.name === 'Line' || runtime.name === 'Name') {
						if (explicit) {
							push('undeclaredVariable', `Variable not defined: '${tok.rawText}'. It opens the ${runtime.name} statement, which gives no value. Declare a variable of that name, or remove Option Explicit.`, at);
						}
						return;
					}
					const required = runtimeRequiredCount(runtime);
					if (required > 0) {
						push('argumentCount', `'${tok.rawText}' needs ${required === 1 ? 'an argument' : `${required} arguments`}, and is named here with none where a value goes. This is a VBE compile error: Argument not optional.`, at);
					} else if (runtime.kind === 'statement') {
						push('subUsedAsValue', `'${tok.rawText}' is a statement, which returns nothing, so it cannot be used as a value. This is a VBE compile error: Expected Function or variable.`, at);
					}
				});
			}
		}, activity);
	}
}

/** How many arguments a library procedure needs: those not in brackets, from its signature. */
function runtimeRequiredCount(runtime: VbaRuntimeFunction): number {
	if (runtime.params) {
		return runtime.params.filter((param) => !param.optional && !param.paramArray).length;
	}
	const open = runtime.signature.indexOf('(');
	const list = open >= 0
		? runtime.signature.slice(open + 1, runtime.signature.indexOf(')', open))
		: runtime.signature.slice(runtime.name.length);
	return list.split(',').map((part) => part.trim()).filter((part) => part !== '' && !part.startsWith('[') && !/^(ParamArray|Optional)\b/i.test(part)).length;
}

/** Per-statement rule: rides the shared procedure-statement walk (audit #0). */
export function checkMemberNotFound(
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	return (procedure) => {
		// Names a bare control reference would lose to: the procedure's
		// parameters and locals.
		const ownNames = new Set(procedure.params.map((param) => param.name.toLowerCase()));
		forEachVariableGroup(procedure.body, (group) => {
			for (const decl of group.declarations) {
				ownNames.add(decl.name.toLowerCase());
			}
		});
		return (stmt) => {
		for (const ref of memberAccessReferences(source, stmt.span)) {
			const surface = resolveMemberPresenceSurfaceAt(
				source,
				ref.dotEndOffset,
				memberCtx,
			);
			const hasMember = surface?.hasMember(ref.member) === true;
			if (!surface?.exhaustive || hasMember) {
				const form = projectMemberFormProblem(source, ref, memberCtx);
				if (form) {
					push('argumentCount', form, ref.memberSpan);
					continue;
				}
				const control = formControlWithoutMember(source, ref, memberCtx, ownNames);
				if (control) {
					push(
						'memberNotFound',
						`Method or data member not found: '${control.name}.${ref.member}'. The form's ${control.type.replace(/^MSForms\./, '')} has no member of that name.`,
						ref.memberSpan,
					);
					continue;
				}
				// A known public member cannot be private, even on a partial surface.
				const owner = hasMember ? undefined
					: privateMemberOwnerAt(source, ref.dotEndOffset, ref.member, memberCtx);
				if (owner) {
					push(
						'memberNotFound',
						`Method or data member not found: '${owner}.${ref.member}'. It is Private to ${owner}, and no reference through an object reaches a Private member, Me included.`,
						ref.memberSpan,
					);
				}
				continue;
			}
			push(
				'memberNotFound',
				`Method or data member not found: '${surface.owner}.${ref.member}'.`,
				ref.memberSpan,
			);
		}
		};
	};
}

interface MemberAccessReference {
	member: string;
	memberSpan: Span;
	dotEndOffset: number;
	/** The statement's tokens, and where the member is among them. */
	toks: readonly VbaToken[];
	index: number;
}

/**
 * A class member used in a form the VBE refuses while compiling (issue #224,
 * measured in Excel 16.0):
 *
 *  - A property whose Get takes a required index, used without one:
 *    `Main = c.Idx`, `c.Idx = 5`. "Argument not optional".
 *  - A Public field of a value type given arguments: `c.Field(1)` is "Wrong
 *    number of arguments or invalid property assignment", and as the target
 *    of a Let, `c.Field(1) = 5`, "Can't assign to read-only property". A
 *    Variant, Collection or Object field takes them, and `c.Field()` compiles.
 */
function projectMemberFormProblem(source: string, ref: MemberAccessReference, memberCtx: MemberCompletionContext): string | undefined {
	const next = ref.toks[ref.index + 1];
	if (next?.rawText === '.' || tokenText(ref.toks[0]) === 'set') {
		return undefined;
	}
	const member = projectClassMemberAt(source, ref.dotEndOffset, ref.member, memberCtx);
	if (!member || member.kind !== 'property') {
		return undefined;
	}
	if (member.signature !== undefined) {
		const open = member.signature.indexOf('(');
		const firstParam = open >= 0 ? member.signature.slice(open + 1).trimStart() : '';
		const indexRequired = firstParam.length > 0 && !firstParam.startsWith(')') && !firstParam.startsWith('[');
		return indexRequired && next?.rawText !== '('
			? `Argument not optional: property '${member.name}' takes an index, as in ${member.signature}. This is a VBE compile error.`
			: undefined;
	}
	const field = member.writable === true && !member.letAccessor && !member.setAccessor;
	const type = normalizeType(member.returns);
	if (!field || type === undefined || type === 'variant' || !isKnownScalarType(type) || next?.rawText !== '(') {
		return undefined;
	}
	const close = matchParenFrom(ref.toks, ref.index + 1);
	if (close !== ref.index + 2 && close > 0) {
		const assigned = ref.toks[close + 1]?.rawText === '=' && ref.index === (ref.toks[0]?.rawText === '.' ? 1 : 2);
		return assigned
			? `'${member.name}' is a field of type ${member.returns}, which takes no index, so '${member.name}(...)' is no place to assign. This is a VBE compile error: Can't assign to read-only property.`
			: `'${member.name}' is a field of type ${member.returns}, which takes no arguments. This is a VBE compile error: Wrong number of arguments or invalid property assignment.`;
	}
	return undefined;
}

/**
 * A form's control reached through the form, with a member its class lacks:
 * `f.T1.Nope`, `Me.T1.Nope`, a bare `T1.Nope` inside the form. The VBE binds
 * those while compiling for the classes in MSFORMS_FORM_CONTROL_MEMBERS; a
 * variable declared As MSForms.TextBox, and a Frame or an OptionButton on the
 * form, it leaves to run time (issue #226, measured in Excel 16.0).
 */
function formControlWithoutMember(
	source: string,
	ref: MemberAccessReference,
	memberCtx: MemberCompletionContext,
	ownNames: ReadonlySet<string>,
): { name: string; type: string } | undefined {
	const controlToken = ref.toks[ref.index - 2];
	const name = controlToken ? tokenName(controlToken) : undefined;
	if (!name || ref.toks[ref.index - 1]?.rawText !== '.') {
		return undefined;
	}
	const lower = name.toLowerCase();
	let type: string | undefined;
	if (ref.toks[ref.index - 3]?.rawText === '.') {
		// `f.T1.Nope`: T1 must be a control of the form the receiver is.
		const form = projectTypeAt(source, ref.dotEndOffset - (ref.toks[ref.index - 1].end - ref.toks[ref.index - 3].end), memberCtx);
		if (form?.kind !== 'userform' || form.exhaustive !== true) {
			return undefined;
		}
		type = form.members.find((member) => member.name.toLowerCase() === lower && /^MSForms\./i.test(member.returns ?? ''))?.returns;
	} else {
		// A bare `T1.Nope` inside the form, where no local or parameter
		// takes the name.
		if (ownNames.has(lower) || memberCtx.meProjectType === undefined) {
			return undefined;
		}
		const self = (memberCtx.projectClassMembers ?? []).find((candidate) => candidate.kind === 'userform'
			&& candidate.exhaustive === true && candidate.name.toLowerCase() === memberCtx.meProjectType?.toLowerCase());
		type = self?.members.find((member) => member.name.toLowerCase() === lower && /^MSForms\./i.test(member.returns ?? ''))?.returns;
	}
	const members = type ? MSFORMS_FORM_CONTROL_MEMBERS[type] : undefined;
	if (!members || members.some((member) => member.toLowerCase() === ref.member.toLowerCase())) {
		return undefined;
	}
	return { name, type: type! };
}

function memberAccessReferences(
	source: string,
	span: Span,
): MemberAccessReference[] {
	const toks = statementTokens(source, span);
	const out: MemberAccessReference[] = [];
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
			toks,
			index: i + 1,
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
	hostModel: HostObjectModel | undefined,
	designerClass: string | undefined,
	push: PushFn,
	projectTypes?: readonly VbaProjectClassMembers[],
	ownMembers: ReadonlySet<string> = new Set(),
): ProcedureStatementVisitor {
	// The host injects Application's members into the global scope, so a bare
	// call may legitimately bind to one of them (Calculate, Volatile, ...).
	const appMembers = applicationMemberNames(hostModel);
	// A module IS its designer's class, so that class's own methods are in
	// scope unqualified: PropertyChanged in a UserControl, Show in a form.
	const designerMembers = designerClassMemberNames(designerClass, hostModel);

	// An Event is no procedure: `Done` or `Call Done(1)` in the class that
	// declares only the Event is "Sub or Function not defined" (issue #266,
	// measured in Excel 16.0). A Sub of the same name, here or public
	// elsewhere, or a local, still binds it.
	const moduleKinds = new Map<string, Set<string>>();
	for (const symbol of symbols.root.children ?? []) {
		const lower = symbol.name.toLowerCase();
		moduleKinds.set(lower, (moduleKinds.get(lower) ?? new Set()).add(symbol.kind));
	}
	const eventOnly = (lower: string): boolean => [...(moduleKinds.get(lower) ?? [])].every((kind) => kind === 'event') && moduleKinds.has(lower);
	const bound = (name: string, procSym: VbaSymbol | undefined): boolean => {
		const lower = name.toLowerCase();
		if (!eventOnly(lower)) {
			return sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, 'call');
		}
		return (procSym?.children ?? []).some((child) => child.name.toLowerCase() === lower)
			|| (projectVisibleSymbols ?? []).some((symbol) => symbol.kind !== 'event' && symbol.name.toLowerCase() === lower);
	};

	const isKnown = (name: string, procSym: VbaSymbol | undefined): boolean => {
		const lower = name.toLowerCase();
		return (
			knownProcedures.has(lower) ||
			bound(name, procSym) ||
			appMembers.has(lower) ||
			designerMembers.has(lower) ||
			ownMembers.has(lower) ||
			resolveHostGlobal(name, hostModel) !== undefined ||
			// The host's hidden Global interface is bare-callable too (issue #34).
			resolveHostGlobalMember(name, hostModel) !== undefined ||
			resolveHostEnum(name, hostModel) !== undefined ||
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
				// A module's name alone is not a call: `Foo` with a module named
				// Foo is "Expected variable or procedure, not module" (issue
				// #125, measured in Excel 16.0).
				const namesModule = (projectTypes ?? []).some(
					(type) => type.kind === 'standardModule' && type.name.toLowerCase() === hit.name.toLowerCase(),
				);
				push(
					'unknownCallStatement',
					namesModule
						? `'${hit.name}' is a module, not a procedure: name the procedure to call, as in '${hit.name}.Bar'. This is a VBE compile error: Expected variable or procedure, not module.`
						: `Sub or Function not defined: '${hit.name}'.`,
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
	implicitMembers: readonly { name: string; type: string }[] | undefined,
	moduleKind: ModuleSymbolKind | undefined,
	hostModel: HostObjectModel | undefined,
	designerClass: string | undefined,
	referencedHosts: readonly string[] | undefined,
	push: PushFn,
	ownMembers: ReadonlySet<string> = new Set(),
): void {
	if (!hasOptionExplicit(mod, activity) || !knownIdentifiers) {
		return;
	}
	// A library name or the project name qualifies a global in an expression
	// as it does in an As clause: `Set app = Excel.Application`,
	// `Word.Application`, `VBAProject.Module2.Twice(4)` (issue #101; each
	// runs in its host). The libraries are the host's own, those merged into
	// its model (Office, MSForms), the ones the project references, and VBA,
	// which isKnown already accepts.
	const libraryQualifiers = libraryQualifierNames(hostModel, referencedHosts);
	// A form's controls are declared by its DESIGNER, not its text, and the seed
	// has three states, not two: a list, a vouched-for-empty list, and no answer
	// at all. Reading no answer as an empty one claimed every control the form's
	// own code-behind names was undeclared - and the VBE stops handing out a
	// designer once the form has been shown, so running your own form turned its
	// code red against source that had just compiled. A form vouched for as EMPTY
	// still reports: that is the case worth keeping, and the member rule draws the
	// same line (issue #48). An Access form or report answers with the list
	// its TypeInfo stream holds, record-source fields included, and the VBE
	// checks a bare name against that list the same way (issue #206).
	if (moduleKind === 'userform' && implicitMembers === undefined) {
		return;
	}
	const implicitMemberNames = new Set(
		(implicitMembers ?? []).map((member) => member.name.toLowerCase()),
	);

	const bracketNamesEvaluate = hostEvaluatesBracketedNames(hostModel);
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const appMembers = applicationMemberNames(hostModel);
	// The designer's class contributes members the text never declares, and a
	// bare reference to one - `Caption = "x"` in a form - is correct code.
	const designerMembers = designerClassMemberNames(designerClass, hostModel);
	// `ReDim items(2) As Long` at procedure level DECLARES `items` when nothing
	// else does (MS-VBAL 5.4.3.3), and Option Explicit accepts it (issue #99,
	// runs in Excel 16.0). Per procedure, below.
	let redimDeclared: ReadonlySet<string> = new Set();
	const isKnown = (
		name: string,
		procSym: VbaSymbol | undefined,
		context: BareIdentifierContext,
	): boolean => {
		const lower = name.toLowerCase();
		return (
			lower === 'vba' ||
			libraryQualifiers.has(lower) ||
			redimDeclared.has(lower) ||
			// A UserForm's controls are members the designer declared, not the
			// module's text; referring to one is correct VBA.
			implicitMemberNames.has(lower) ||
			designerMembers.has(lower) ||
			// The module's own object: UsedRange in a sheet, Tag in a form
			// (issue #228).
			ownMembers.has(lower) ||
			sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, context) ||
			knownIdentifiers.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name, hostModel) !== undefined ||
			// Members of the host's hidden Global interface are callable bare
			// (Word's InchesToPoints, Excel's Union) - issue #34.
			resolveHostGlobalMember(name, hostModel) !== undefined ||
			resolveHostConstant(name, hostModel) !== undefined ||
			// An enum name is a legal qualifier: `XlAxisType.xlCategory` is
			// ordinary VBA, and Option Explicit called the qualifier undeclared.
			resolveHostEnum(name, hostModel) !== undefined ||
			resolveRuntimeConstant(name) !== undefined ||
			resolveRuntimeObject(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined ||
			// VBA's own enums and modules qualify their members the same way:
			// `VbMsgBoxResult.vbYes`, `ColorConstants.vbRed`, `Strings.Left`.
			resolveVbaLibraryQualifier(name) !== undefined
		);
	};

	let eol: string | undefined;
	const moduleEol = (): string => eol ??= detectEol(source);
	const ctxForTypes = {
		moduleName: symbols.moduleName,
		moduleKind,
		model: hostModel,
		projectClassMembers: projectMembers,
		projectVisibleSymbols,
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = procedureSymbolFor(symbols, member);
		redimDeclared = redimTargetNamesIn(source, member.body, activity);
		const declarationData = declarationDataFor(source, member, moduleEol);
		forEachUndeclaredReferenceSpan(source, member.body, (span) => {
			const reported = new Set<string>();
			const report = (
				name: string,
				span: Span,
				mode: 'assigning to it' | 'using it',
				context: BareIdentifierContext,
				assignedType?: () => string,
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
					declarationData(name, assignedType?.()),
				);
			};
			const scalarTarget = bareAssignmentTarget(source, span);
			const objectTarget = scalarTarget ? undefined : setAssignmentTarget(source, span);
			const target = scalarTarget ?? objectTarget;
			if (target) {
				// A THUNK, called past the isKnown guard inside `report`. Typing
				// the right-hand side costs a bind of the module, and running it
				// for every assignment made a module whose variables are all
				// declared - which is most modules - pay that for no findings
				// (github.com/WilliamSmithEdward/xlide_vscode/issues/62).
				//
				// A `Set` makes the name an object whatever the right-hand side
				// turns out to be; anything else is read off the expression, and
				// Variant where nothing narrows it.
				report(
					target.name,
					target.span,
					'assigning to it',
					'assignmentTarget',
					() => (objectTarget ? 'Object' : assignedValueType(source, span, ctxForTypes)),
				);
			}
			for (const ref of undeclaredReadReferences(
				source,
				span,
				(name) => isKnown(name, procSym, 'expression'),
				moduleSignatures,
				projectMembers,
			)) {
				if (ref.bracketed && bracketNamesEvaluate) {
					continue;
				}
				report(ref.name, ref.span, 'using it', 'expression');
			}
		}, activity);
	}

	// A Const's value and an Enum member's value name things too: `Const K =
	// asdf` is "Variable not defined", and `eB = asdf` in an Enum "Constant
	// expression required" (issue #369, measured in Excel 16.0).
	const checkValue = (span: Span, procSym: VbaSymbol | undefined, enumMember: boolean, what = "a Const's value"): void => {
		// The names in each value, after its `=`. A value calls nothing, so
		// every name in it not after a `.` is read; a declaration's own names
		// stand before an `=`.
		const toks = statementTokens(source, span);
		let inValue = false;
		for (let i = 0; i < toks.length; i++) {
			const tok = toks[i];
			if (tok.rawText === '=') {
				inValue = true;
				continue;
			}
			if (tok.rawText === ',') {
				inValue = false;
				continue;
			}
			const name = tok.kind === 'identifier' ? tokenName(tok) : undefined;
			// A name before a `.` qualifies: `Module2.B1`, `Excel.xlUp`.
			const qualifier = toks[i + 1]?.rawText === '.';
			if (!inValue || !name || qualifier || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!' || isKnown(name, procSym, 'expression')) {
				continue;
			}
			push(
				'undeclaredVariable',
				enumMember
					? `'${name}' is not defined, and an Enum member's value must be a constant. This is a VBE compile error: Constant expression required.`
					: `Variable not defined: '${name}'. Declare it before using it in ${what}, or remove Option Explicit.`,
				{ start: span.start + tok.start, end: span.start + tok.end },
			);
		}
	};
	redimDeclared = new Set();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && member.isConst) {
			checkValue(member.span, undefined, false);
		} else if (member.kind === 'Enum') {
			for (const item of member.members) {
				if (item.valueRaw !== undefined && !activity?.isInactive(item.span)) {
					checkValue(item.span, undefined, true);
				}
			}
		} else if (member.kind === 'Procedure') {
			// An Optional parameter's default is a constant from outside the
			// procedure: `Optional x As Long = y` with nothing named y is
			// "Variable not defined" (issue #445, measured in Excel 16.0).
			for (const param of member.params) {
				if (param.defaultRaw !== undefined) {
					checkValue(param.span, undefined, false, "an Optional parameter's default");
				}
			}
			const procSym = procedureSymbolFor(symbols, member);
			forEachVariableGroup(member.body, (group) => {
				if (group.isConst) {
					checkValue(group.span, procSym, false);
				}
			}, activity);
		}
	}
}

/**
 * Whether `[name]` on its own is a HOST LOOKUP rather than a variable.
 *
 * In Excel the square brackets are shorthand for `Application.Evaluate`, so
 * `[A1]` and `[TaxRate]` are ordinary code that compiles and needs no
 * declaration. Word has no such feature - measured in the VBE, `v = [foo]`
 * with nothing declaring `foo` is a compile error there - so the report is
 * right for Word and PowerPoint and must stay.
 *
 * Only a POSITIVELY identified non-Excel host reports. An absent model is
 * Excel's by default (issue #28), and a host whose model knows nothing
 * asserts nothing, so both of those suppress rather than guess.
 */
function hostEvaluatesBracketedNames(hostModel: HostObjectModel | undefined): boolean {
	return hostModel?.hostName === undefined || hostModel.hostName === 'Excel';
}

/**
 * Lower-cased names a procedure's ReDim statements size: `ReDim name(...)`,
 * `ReDim Preserve name(...)`, and each further `, name(...)`. A ReDim of an
 * undeclared name declares it, so these count as declared for the procedure.
 */
function redimTargetNamesIn(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Set<string> {
	const out = new Set<string>();
	const visit = (nodes: readonly BodyNode[]): void => {
		for (const node of nodes) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (isLeafStatement(node)) {
				for (const span of statementAndBranchSpans(node)) {
					const toks = statementTokens(source, span);
					let i = firstExecutableTokenIndex(toks);
					if (tokenText(toks[i]) !== 'redim') {
						continue;
					}
					i++;
					if (tokenText(toks[i]) === 'preserve') {
						i++;
					}
					let depth = 0;
					for (let k = i; k < toks.length; k++) {
						const raw = toks[k].rawText;
						if (raw === '(') {
							depth++;
						} else if (raw === ')') {
							depth--;
						} else if (depth === 0 && (k === i || toks[k - 1].rawText === ',')) {
							const name = tokenName(toks[k]);
							if (name && toks[k + 1]?.rawText === '(') {
								out.add(name.toLowerCase());
							}
						}
					}
				}
				continue;
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
		}
	};
	visit(body);
	return out;
}

/**
 * Lower-cased names that may qualify a global in an expression: the libraries
 * whose types the host model carries (its own and the shared ones merged into
 * it), the libraries the project references, and the project itself, which is
 * `VBAProject` unless renamed. An absent model is Excel's by default.
 */
const libraryQualifierModels = new WeakMap<HostObjectModel, {
	types: HostObjectModel['types']; enums: HostObjectModel['enums']; names: ReadonlySet<string>;
}>();

function libraryQualifierNames(
	hostModel: HostObjectModel | undefined,
	referencedHosts: readonly string[] | undefined,
): Set<string> {
	const model = hostModel ?? getExcelObjectModel();
	let cached = libraryQualifierModels.get(model);
	if (!cached || cached.types !== model.types || cached.enums !== model.enums) {
		const names = new Set<string>();
		for (const qualified of Object.keys(model.types)) {
			const dot = qualified.indexOf('.');
			if (dot > 0) { names.add(qualified.slice(0, dot).toLowerCase()); }
		}
		for (const enumeration of Object.values(model.enums ?? {})) {
			if (enumeration.library) { names.add(enumeration.library.toLowerCase()); }
		}
		cached = { types: model.types, enums: model.enums, names };
		libraryQualifierModels.set(model, cached);
	}
	const out = new Set<string>(['vbaproject', ...cached.names]);
	for (const token of referencedHosts ?? []) {
		const name = HOST_LIBRARY_NAMES[token as VbaHostToken];
		if (name) {
			out.add(name.toLowerCase());
		}
	}
	return out;
}

function undeclaredReadReferences(
	source: string,
	span: Span,
	isKnown: (name: string) => boolean,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): Array<{ name: string; span: Span; bracketed: boolean }> {
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

/**
 * The type the right-hand side of an assignment gives the name being declared.
 * `Variant` where nothing narrows it, which is what VBA would have given the
 * name anyway and what the developer would have typed by hand.
 */
function assignedValueType(
	source: string,
	span: Span,
	ctx: ExpressionTypeContext,
): string {
	const target = bareAssignmentTarget(source, span);
	const tokens = target?.valueTokens.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	if (!tokens || tokens.length === 0) {
		return 'Variant';
	}
	const valueSpan = {
		start: span.start + tokens[0].start,
		end: span.start + tokens[tokens.length - 1].end,
	};
	const value = resolveExpressionType(source, valueSpan, ctx);
	return value?.complete ? value.type : 'Variant';
}

/**
 * A `Dim` for the missing name, placed at the top of the enclosing procedure
 * under any declarations already there - where a VBA author puts them, and
 * where `Dim` must appear before the first use anyway.
 *
 * Only offered for an ASSIGNMENT to the name: a bare read gives nothing to
 * infer a type from, and declaring a name the author only reads is as likely to
 * be papering over a typo as to be the fix.
 */
function declarationDataFor(
	source: string,
	member: ProcedureNode,
	moduleEol: () => string,
): (name: string, declaredType: string | undefined) => VbaDiagnosticData | undefined {
	// The insertion site depends only on this procedure and source. Resolve it
	// once, and only when a diagnostic actually offers a declaration edit.
	let site: { insertAt: number; indent: string; eol: string } | null | undefined;
	return (name, declaredType) => {
		// Any letter the code page holds starts an identifier (issue #207).
		if (!declaredType || !/^\p{L}[\p{L}\p{N}_]*$/u.test(name)) {
			return undefined;
		}
		if (site === undefined) {
			const insertAt = declarationInsertOffset(source, member);
			site = insertAt === undefined ? null : {
				insertAt,
				indent: leadingWhitespaceOfLineAt(source, insertAt),
				eol: moduleEol(),
			};
		}
		if (!site) {
			return undefined;
		}
		return {
			declareVariable: {
				variableName: name,
				declaredType,
				edit: {
					span: { start: site.insertAt, end: site.insertAt },
					newText: `${site.indent}Dim ${name} As ${declaredType}${site.eol}`,
				},
			},
		};
	};
}

/** Offset of the line after the procedure's leading declarations. */
function declarationInsertOffset(source: string, member: ProcedureNode): number | undefined {
	let insertAt: number | undefined;
	for (const stmt of member.body) {
		if (!isLeafStatement(stmt)) {
			break;
		}
		const text = source.slice(stmt.span.start, stmt.span.end).trim();
		if (text === '' || text.startsWith("'")) {
			continue;
		}
		if (!/^(?:Dim|Static|Const)\b/i.test(text)) {
			break;
		}
		insertAt = lineStartAfter(source, stmt.span.end);
	}
	return insertAt ?? firstBodyLineStart(source, member);
}

/** Start of the line following `offset`. */
function lineStartAfter(source: string, offset: number): number {
	const next = lineEndAtOrAfter(source, offset);
	return next < source.length ? next + (source[next] === '\r' && source[next + 1] === '\n' ? 2 : 1) : next;
}

/** Start of the first body line of a procedure, just after its header line. */
function firstBodyLineStart(source: string, member: ProcedureNode): number | undefined {
	const first = member.body.find((stmt) => isLeafStatement(stmt) || 'body' in stmt);
	if (first) {
		return lineStartAtAnyBreak(source, first.span.start);
	}
	return lineStartAfter(source, member.span.start);
}

/** The indentation of the line at `offset`, reused for the inserted line. */
function leadingWhitespaceOfLineAt(source: string, offset: number): string {
	const start = lineStartAtAnyBreak(source, offset);
	const end = lineEndAtOrAfter(source, start);
	const line = source.slice(start, end);
	return /^[ \t]*/.exec(line)?.[0] ?? '';
}
