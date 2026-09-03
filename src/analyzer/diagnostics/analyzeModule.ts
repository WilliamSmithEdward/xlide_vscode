// Active VBA diagnostics engine (MS-VBAL Phase 5).
//
// `analyzeModule` runs the high-confidence semantic rules from the rule
// catalogue over one module's source and returns offset-based diagnostics. It
// is pure (no `vscode`): the editor layer converts spans to ranges and severity
// names to the VS Code enum. Structural block-balance checking stays in
// `src/vbaStructuralDiagnostics.ts` (analyzeVbaStructure); this engine adds the semantic rules on
// top, so the two do not overlap or double-report.
//
// Design rule (see /memories): no "looks like" heuristics. Every rule here is
// deterministic - it flags a construct only when the language guarantees it is
// an error. The one cross-module rule, `unknown-call`, fires only on a call
// statement whose callee is a bare (non-member) identifier - the lone-identifier
// form, the parenless-argument form (`MsgBox "hi"`), or `Call name` - whose name
// resolves to no procedure anywhere in the project, no VBA runtime
// function/statement, no host global or Application member, and no in-scope
// declaration - the unambiguous VBE "Sub or Function not defined" error.
// The `undeclared-variable` rule follows the same conservative pattern: it runs
// only when the caller has supplied project-visible names, and it scans
// statement-level reads while deliberately skipping unresolved external-style
// calls and type-name positions to avoid false positives.

import { tokenizeCached } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import {
	absoluteSpan,
	physicalLineSpanAtOffset,
	statementTokens,
	walkProcedureStatements,
	type ProcedureStatementVisitor,
} from './walker';
import {
	walkProcedureExpressions,
	type ProcedureExpressionVisitor,
} from './exprWalk';
import type {
	ModuleNode,
	ProcedureNode,
	Span,
} from '../parser/nodes';
import { hostObjectModelForToken } from '../host/hostRegistry';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { ModuleSymbolKind } from '../symbols/symbolModel';
import { type MemberCompletionContext } from '../completion/memberAccess';
import {
	DIAGNOSTIC_RULES,
	DiagnosticRuleName,
	DiagnosticSeverity,
	normalizeDiagnosticSeverityOverride,
} from './ruleMetadata';

export type {
	AnalyzeModuleOptions,
	DiagnosticSeverityOverrides,
	VbaCreateProcedureStubData,
	VbaDiagnostic,
	VbaDiagnosticData,
	VbaMissingRequiredArgumentPlaceholderData,
} from './analysisContext';
import { isObjectModuleKind } from './analysisContext';
import type {
	AnalyzeModuleOptions,
	DiagnosticSeverityOverrides,
	PushFn,
	RulePassContext,
	VbaDiagnostic,
	VbaDiagnosticData,
} from './analysisContext';
import { DIAGNOSTIC_RULE_REGISTRY } from './registry';
import {
	incompleteMemberAccess,
	isNonUnaryBinaryOperator,
} from './rules/expressions';

/** Resolves the effective severity of a rule, or undefined when switched off. */
function severityOf(
	rule: DiagnosticRuleName,
	overrides: DiagnosticSeverityOverrides | undefined,
): DiagnosticSeverity | undefined {
	const meta = DIAGNOSTIC_RULES[rule];
	const override = normalizeDiagnosticSeverityOverride(meta.code, overrides?.[meta.code]);
	if (override === 'off') {
		return undefined;
	}
	return override ?? meta.defaultSeverity;
}

/**
 * Analyzes one VBA module source and returns its active diagnostics.
 * Never throws: any internal failure yields an empty list.
 */
export function analyzeModule(
	source: string,
	opts: AnalyzeModuleOptions = {},
): VbaDiagnostic[] {
	try {
		return runRules(source, withResolvedHostModel(opts));
	} catch {
		return [];
	}
}

/**
 * Resolves the `host` token into a hostModel once, up front, so every
 * `opts.hostModel` consumer inherits the caller's choice. An explicit
 * hostModel wins; absent both, the Excel defaults ride as they always have.
 */
export function withResolvedHostModel(opts: AnalyzeModuleOptions): AnalyzeModuleOptions {
	if (opts.hostModel !== undefined || opts.host === undefined) {
		return opts;
	}
	const resolved = hostObjectModelForToken(opts.host);
	return resolved === undefined ? opts : { ...opts, hostModel: resolved };
}

export function incompleteExpressionEditSpan(
	source: string,
	offset: number,
): Span | undefined {
	const line = physicalLineSpanAtOffset(source, offset);
	const statement = activeStatementSpanOnLine(source, line, offset);
	return incompleteMemberAccess(source, statement, { includeLeadingDot: true })?.span
		?? trailingBinaryOperatorEditSpan(source, statement, offset)
		?? unmatchedOpenParenEditSpan(source, statement, offset);
}

function runRules(
	source: string,
	opts: AnalyzeModuleOptions,
): VbaDiagnostic[] {
	const moduleName = opts.moduleName ?? 'Module';
	const moduleKind = opts.moduleKind ?? 'standard';
	const overrides = opts.severityOverrides;

	// Provenance tracking for incremental re-analysis: pushes during the shared
	// procedure walks are stamped with the walked procedure so the incremental
	// layer can cache and splice them per procedure; pushes from module-wide
	// rule passes are recomputed every time. When a walkProcedureFilter is
	// active, walk-phase pushes for filtered-out (clean) procedures are dropped
	// entirely - their cached diagnostics get spliced back by the caller.
	let phase: 'run' | 'walk' = 'run';
	let walkMemberStart: number | undefined;
	let walkMemberIncluded = true;

	const pushInto = (sink: VbaDiagnostic[]): PushFn => (
		rule: DiagnosticRuleName,
		message: string,
		span: Span,
		data?: VbaDiagnosticData,
	): void => {
		if (phase === 'walk' && !walkMemberIncluded) {
			return;
		}
		const severity = severityOf(rule, overrides);
		if (!severity) {
			return;
		}
		const meta = DIAGNOSTIC_RULES[rule];
		sink.push({
			code: meta.code,
			message,
			severity,
			span,
			specReference: meta.specReference,
			origin: phase,
			...(phase === 'walk' && walkMemberStart !== undefined ? { walkMemberStart } : {}),
			...(data ? { data } : {}),
		});
	};

	const mod = opts.parsedModule ?? parseModule(source);
	const ctx: RulePassContext = {
		source,
		moduleName,
		moduleKind,
		opts,
		mod,
		symbols: buildModuleSymbols(moduleName, moduleKind, source, {
			conditionalCompilation: opts.conditionalCompilation,
			parsedModule: mod,
		}),
		activity: createConditionalActivityTracker(mod, opts.conditionalCompilation),
		memberCtx: diagnosticMemberCompletionContext(opts, source, mod),
	};

	// Each rule reports into its own buffer; per-statement rules all ride the
	// one shared procedure-statement walk (audit #0). Flushing the buffers in
	// registry order keeps the historical rule-major diagnostic order.
	const buffers: VbaDiagnostic[][] = [];
	const statementVisitors: ProcedureStatementVisitor[] = [];
	const expressionVisitors: ProcedureExpressionVisitor[] = [];
	for (const rule of DIAGNOSTIC_RULE_REGISTRY) {
		const buffer: VbaDiagnostic[] = [];
		buffers.push(buffer);
		const push = pushInto(buffer);
		// Isolate each rule: one rule throwing during construction or its eager
		// run() must not discard every other rule's diagnostics for the module.
		try {
			if (rule.run) {
				rule.run(ctx, push);
			}
			if (rule.procedureStatements) {
				statementVisitors.push(rule.procedureStatements(ctx, push));
			}
			if (rule.procedureExpressions) {
				expressionVisitors.push(rule.procedureExpressions(ctx, push));
			}
		} catch {
			// Degrade only this rule; keep the rest of the pass intact.
		}
	}
	// A visitor throwing during a shared walk must not blank the run()-based
	// diagnostics already collected, nor the other walk.
	const filter = opts.walkProcedureFilter;
	const walkHooks = {
		beforeMember: (member: ProcedureNode): void => {
			walkMemberStart = member.span.start;
			walkMemberIncluded = filter ? filter(member) : true;
		},
		skipBody: (member: ProcedureNode): boolean => (filter ? !filter(member) : false),
	};
	phase = 'walk';
	try {
		walkProcedureStatements(ctx.mod, ctx.activity, statementVisitors, walkHooks);
	} catch { /* degrade gracefully */ }
	try {
		walkProcedureExpressions(ctx.mod, ctx.activity, expressionVisitors, walkHooks);
	} catch { /* degrade gracefully */ }
	phase = 'run';
	walkMemberStart = undefined;
	walkMemberIncluded = true;

	const out: VbaDiagnostic[] = [];
	for (const buffer of buffers) {
		out.push(...buffer);
	}
	return out;
}

function diagnosticMemberCompletionContext(
	opts: AnalyzeModuleOptions,
	source: string,
	mod: ModuleNode,
): MemberCompletionContext {
	const ctx: MemberCompletionContext = {
		projectClassMembers: opts.projectClassMembers,
		allowSetAssignmentRefinement: false,
		model: opts.hostModel,
		// Audit #1: member resolution used to re-parse the module and re-lex the
		// source prefix once per dotted reference. Hand it the per-pass AST and
		// the shared full-source token stream instead.
		parsedModule: mod,
		sourceTokens: tokenizeCached(source).filter((t) => t.kind !== 'comment'),
	};
	const meProjectType = meProjectTypeFor(opts.moduleName, opts.moduleKind);
	if (meProjectType) {
		ctx.meProjectType = meProjectType;
	}
	const meType = meHostTypeFor(opts.moduleName, opts.moduleKind, opts.host, opts.designerClass);
	if (meType) {
		ctx.meType = meType;
	}
	return ctx;
}

function meProjectTypeFor(
	moduleName: string | undefined,
	moduleKind: ModuleSymbolKind | undefined,
): string | undefined {
	return moduleName && isObjectModuleKind(moduleKind) ? moduleName : undefined;
}

function meHostTypeFor(
	moduleName: string | undefined,
	moduleKind: ModuleSymbolKind | undefined,
	host?: string,
	designerClass?: string,
): string | undefined {
	// A designer class is what the module IS, whatever kind it is listed as: a
	// VB6 form's `Me` reaches Arrange and Show from `VB.MDIForm`, which the
	// module's own text never declares. The module's project type still applies
	// alongside, so `Me` keeps its own procedures too.
	if (designerClass) {
		return designerClass;
	}
	if (!moduleName || moduleKind !== 'document') {
		return undefined;
	}
	const lower = moduleName.toLowerCase();
	switch ((host ?? 'excel').toLowerCase()) {
		case 'excel':
			return lower === 'thisworkbook' ? 'Excel.Workbook' : undefined;
		case 'word':
			return lower === 'thisdocument' ? 'Word.Document' : undefined;
		default:
			// PowerPoint has no document modules; other hosts' document
			// surfaces are unmodelled, and silence beats a wrong type.
			return undefined;
	}
}

function activeStatementSpanOnLine(source: string, line: Span, offset: number): Span {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	const toks = statementTokens(source, line);
	let depth = 0;
	let start = line.start;
	let end = line.end;
	for (const tok of toks) {
		if (tok.kind === 'punctuation') {
			if (tok.rawText === '(') {
				depth++;
			} else if (tok.rawText === ')') {
				depth = Math.max(0, depth - 1);
			}
			continue;
		}
		if (tok.kind !== 'colon' || depth !== 0) {
			continue;
		}
		const colon = absoluteSpan(line, tok);
		if (safeOffset <= colon.start) {
			end = colon.start;
			break;
		}
		start = colon.end;
	}
	return { start, end };
}

function trailingBinaryOperatorEditSpan(
	source: string,
	span: Span,
	offset: number,
): Span | undefined {
	const toks = statementTokens(source, span);
	const last = toks[toks.length - 1];
	if (!isNonUnaryBinaryOperator(last)) {
		return undefined;
	}
	const active = absoluteSpan(span, last);
	if (offset < active.start) {
		return undefined;
	}
	const cursorTail = source.slice(active.end, Math.min(offset, span.end));
	return /^[ \t]*$/.test(cursorTail) ? active : undefined;
}

function unmatchedOpenParenEditSpan(
	source: string,
	span: Span,
	offset: number,
): Span | undefined {
	const stack: VbaToken[] = [];
	for (const tok of statementTokens(source, span)) {
		if (tok.kind !== 'punctuation') {
			continue;
		}
		if (tok.rawText === '(') {
			stack.push(tok);
		} else if (tok.rawText === ')' && stack.length > 0) {
			stack.pop();
		}
	}
	const firstUnmatched = stack[0];
	if (!firstUnmatched) {
		return undefined;
	}
	const active = absoluteSpan(span, firstUnmatched);
	return offset >= active.start ? active : undefined;
}
