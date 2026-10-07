// Incremental rule re-analysis: when only procedure BODIES changed between two
// versions of a module, re-run the expensive per-procedure statement/expression
// walks for the dirty procedures only and splice the cached walk diagnostics of
// the untouched ones (offset-shifted). Eager rules that support filtering also
// reuse procedure findings, while module facts still run over the full AST.
// Ordinary declaration/signature changes invalidate their consumers. Directives,
// implicit default members, indirect declarations and external context retain full passes.

import { statementTokens, tokenName } from './walker';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { writtenNamesIn } from './moduleState';
import { analyzeModule } from './analyzeModule';
import type { AnalyzeModuleOptions, VbaDiagnostic } from './analysisContext';
import type { ModuleNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import { procedureSymbolFor } from './analysisContext';
import { parseModule } from '../parser/parseModule';
import { calleeModuleVariableEffects, normalizeType, typeEnvironmentFor } from './typeInference';
import { calleeKeepsArgument, calleeMemberCallFingerprint } from './calleeArguments';
import { checkAnalysisCancellation } from './analysisCancellation';

interface ProcedureRecord {
	key: string;
	/** Procedure fullSpan in the source it was recorded against. */
	start: number;
	end: number;
	/** Complete parsed body interior, with conservative fallback for recovery nodes. */
	bodyStart: number;
	bodyEnd: number;
	bodyText: string;
	effectText: string;
	/** Name dependencies are immutable and reusable when body text is unchanged. */
	references?: ReadonlySet<string>;
}

export interface ModuleRulesIncrementalState {
	/** Previous immutable context for comparing facts that actually propagate. */
	source: string;
	symbols: ReturnType<typeof buildModuleSymbols>;
	envelope: string;
	writtenNames: ReadonlySet<string>;
	procedures: ProcedureRecord[];
	/** Walk-origin diagnostics per procedure (same index as `procedures`). */
	walkDiagnosticsByProcedure: VbaDiagnostic[][];
	fingerprint: readonly unknown[];
}

export interface ModuleRulesIncrementalResult {
	diagnostics: VbaDiagnostic[];
	state: ModuleRulesIncrementalState;
	mode: 'full' | 'incremental';
}

/** Small modules can recheck four bodies; large modules fall back above half. */
const MIN_DIRTY_PROCEDURES = 4;

function procedureMembers(mod: ModuleNode): ProcedureNode[] {
	const out: ProcedureNode[] = [];
	for (const member of mod.members) {
		if (member.kind === 'Procedure') {
			out.push(member);
		}
	}
	return out;
}

function recordFor(member: ProcedureNode, source: string, shadows: ReadonlySet<string>): ProcedureRecord {
	const first = member.body[0];
	const last = member.body[member.body.length - 1];
	const bodyStart = member.bodySpan?.start ?? first?.span.start ?? member.span.start;
	const bodyEnd = member.bodySpan?.end ?? last?.span.end ?? member.span.start;
	return {
		key: `${member.procKind}:${member.name.toLowerCase()}`,
		start: member.span.start,
		end: member.span.end,
		bodyStart,
		bodyEnd,
		bodyText: source.slice(bodyStart, bodyEnd),
		effectText: procedureEffects(member, source, shadows),
	};
}

/**
 * Literal-only output cannot change return expressions or argument values.
 * Module-variable effects and member replay facts are compared separately. Keep
 * every other statement, declaration and block in the dependency fingerprint.
 * Use parsed statement spans: comments, line continuations and labels cannot
 * accidentally turn part of a different statement into disposable output.
 */
function procedureEffects(member: ProcedureNode, source: string, shadows: ReadonlySet<string>): string {
	const spans = member.body.filter(node => isLeafStatement(node) && !(node.kind === 'Statement' && node.singleLineIfBranches))
		.map(node => node.span).filter(span => {
			const text = source.slice(span.start, span.end);
			const head = /^\s*(Debug|MsgBox)\b/i.exec(text)?.[1].toLowerCase();
			if (!head || shadows.has(head)) { return false; }
			return /^\s*(?:Debug\s*\.\s*Print|MsgBox)\b\s*(?:"(?:[^"\r\n]|"")*"|\d+(?:\.\d+)?|[ \t()+*/&^=<>,-])*\s*$/i.test(text);
		});
	const parts: string[] = [];
	let pos = member.body[0]?.span.start ?? member.span.start;
	for (const span of spans) { parts.push(source.slice(pos, span.start)); pos = span.end; }
	parts.push(source.slice(pos, member.body.at(-1)?.span.end ?? pos));
	// Dropping an inert statement can leave an extra blank line in the body.
	return parts.join('').split(/\r?\n/).filter(line => line.trim().length > 0).join('\n');
}

/**
 * Everything outside procedure bodies, with each body replaced by a marker.
 * Two sources with equal envelopes have identical declaration sections,
 * procedure signatures, inter-procedure text, and directive structure - the
 * only differences are inside procedure bodies.
 */
function envelopeOf(source: string, records: readonly ProcedureRecord[]): string {
	const parts: string[] = [];
	let pos = 0;
	for (const record of records) {
		parts.push(source.slice(pos, record.bodyStart), '\x00');
		pos = record.bodyEnd;
	}
	parts.push(source.slice(pos));
	return parts.join('');
}

function sameFingerprint(a: readonly unknown[], b: readonly unknown[]): boolean {
	return a.length === b.length && a.every((value, i) => Object.is(value, b[i]));
}

interface SurfaceMatch {
	previousIndices: number[];
	changedHeaders: Set<number>;
	changedClosers: Set<number>;
	changedNames: Set<string>;
}

/** Keep unknown module constructs, options and directives in the strict envelope. */
function declarationEnvelope(source: string, mod: ModuleNode): string {
	const parts: string[] = [];
	let pos = 0;
	for (const member of mod.members) {
		if (member.kind !== 'Procedure' && member.kind !== 'VariableGroup') { continue; }
		parts.push(source.slice(pos, member.span.start));
		pos = member.span.end;
	}
	parts.push(source.slice(pos));
	return parts.join('').split(/\r?\n/).filter(line => line.trim().length > 0).join('\n');
}

/** Match unchanged bodies across ordinary declaration edits, additions and removals. */
function matchDeclarationSurface(source: string, mod: ModuleNode, records: ProcedureRecord[], prev: ModuleRulesIncrementalState): SurfaceMatch | undefined {
	const oldMod = parseModule(prev.source);
	if ([mod, oldMod].some(module => module.members.some(member => member.kind === 'Procedure' && !member.bodySpan))
		|| declarationEnvelope(source, mod) !== declarationEnvelope(prev.source, oldMod)) { return undefined; }
	const oldIndices = new Map(prev.procedures.map((record, i) => [record.key, i]));
	if (oldIndices.size !== prev.procedures.length || new Set(records.map(record => record.key)).size !== records.length) { return undefined; }
	const previousIndices = records.map(record => oldIndices.get(record.key) ?? -1);
	const currentMembers = procedureMembers(mod);
	const oldMembers = procedureMembers(oldMod);
	const attributes = (text: string, member: ProcedureNode | undefined) => member?.attributes?.map(attribute => text.slice(attribute.span.start, attribute.span.end)).join('\n') ?? '';
	for (let i = 0; i < records.length; i++) {
		// Default-member attributes affect implicit calls that never spell the name.
		if (attributes(source, currentMembers[i]) !== attributes(prev.source, oldMembers[previousIndices[i]])) { return undefined; }
	}
	const changedHeaders = new Set<number>();
	const changedClosers = new Set<number>();
	const changedNames = new Set<string>();
	const header = (text: string, record: ProcedureRecord) => text.slice(record.start, record.bodyStart);
	for (let i = 0; i < records.length; i++) {
		const old = prev.procedures[previousIndices[i]];
		if (old && source.slice(records[i].bodyEnd, records[i].end) !== prev.source.slice(old.bodyEnd, old.end)) { changedClosers.add(i); }
		if (!old || header(source, records[i]) !== header(prev.source, old)) {
			if (currentMembers[i].attributes?.length) { return undefined; }
			changedHeaders.add(i);
			changedNames.add(records[i].key.slice(records[i].key.indexOf(':') + 1));
		}
	}
	const retained = new Set(previousIndices);
	for (let i = 0; i < prev.procedures.length; i++) {
		if (!retained.has(i)) {
			if (oldMembers[i].attributes?.length || /^[ \t]*#/m.test(prev.procedures[i].bodyText)) { return undefined; }
			changedNames.add(prev.procedures[i].key.slice(prev.procedures[i].key.indexOf(':') + 1));
		}
	}
	const declarations = (text: string, module: ModuleNode) => {
		const byName = new Map<string, string>();
		for (const member of module.members) {
			if (member.kind !== 'VariableGroup') { continue; }
			for (const declaration of member.declarations) {
				const name = declaration.name.toLowerCase();
				if (byName.has(name)) { return undefined; }
				byName.set(name, text.slice(member.span.start, member.span.end));
			}
		}
		return byName;
	};
	const oldDeclarations = declarations(prev.source, oldMod);
	const newDeclarations = declarations(source, mod);
	if (!oldDeclarations || !newDeclarations) { return undefined; }
	const changedDeclarations = new Set<string>();
	for (const name of new Set([...oldDeclarations.keys(), ...newDeclarations.keys()])) {
		if (oldDeclarations.get(name) !== newDeclarations.get(name)) { changedDeclarations.add(name); changedNames.add(name); }
	}
	if ([...currentMembers, ...oldMembers].some(member => member.attributes?.length && changedNames.has(member.name.toLowerCase()))) { return undefined; }
	// A changed Const may be used in another Const, a UDT bound or an Enum.
	// Those indirect declaration dependencies retain a full pass.
	for (const [text, module] of [[source, mod], [prev.source, oldMod]] as const) {
		for (const member of module.members) {
			if (member.kind === 'Procedure') { continue; }
			if (member.kind === 'VariableGroup' && member.declarations.every(declaration => changedDeclarations.has(declaration.name.toLowerCase()))) { continue; }
			if (statementTokens(text, member.span).some(token => changedDeclarations.has(tokenName(token)?.toLowerCase() ?? ''))) { return undefined; }
		}
	}
	return { previousIndices, changedHeaders, changedClosers, changedNames };
}

/** Bare names bind to a procedure's locals/parameters before module procedures. */
function procedureReferences(source: string, member: ProcedureNode, symbols: ReturnType<typeof buildModuleSymbols>, moduleName: string): ReadonlySet<string> {
	const localNames = new Set((procedureSymbolFor(symbols, member)?.children ?? []).map(symbol => symbol.name.toLowerCase()));
	for (const parameter of member.params) { localNames.add(parameter.name.toLowerCase()); }
	const tokens = statementTokens(source, member.span);
	const declarationNames = new Set([member.nameSpan?.start, ...member.params.map(parameter => parameter.nameSpan?.start)]);
	let types: ReadonlyMap<string, string> | undefined;
	const references = new Set<string>();
	for (let i = 0; i < tokens.length; i++) {
		const token = tokens[i];
		if (token.kind === 'stringLiteral') {
			// Dynamic call targets are names (possibly workbook/module qualified).
			// Words inside prose such as "Invalid value" do not call Value().
			const text = token.rawText.slice(1, -1).replace(/""/g, '"').trim().toLowerCase();
			const name = /(?:^|[.!])([a-z_]\w*)$/.exec(text)?.[1];
			if (name) { references.add(name); }
			continue;
		}
		const name = tokenName(token)?.toLowerCase();
		if (!name || tokens[i + 1]?.rawText === ':=') { continue; }
		const absoluteStart = member.span.start + token.start;
		const inBody = absoluteStart >= (member.bodySpan?.start ?? member.body[0]?.span.start ?? member.span.end);
		if (!inBody && declarationNames.has(absoluteStart)) { continue; }
		const qualified = tokens[i - 1]?.rawText === '.' || tokens[i - 1]?.rawText === '!';
		const typeName = /^(as|new)$/i.test(tokens[i - 1]?.rawText ?? '');
		if (inBody && !qualified && !typeName && localNames.has(name)) { continue; }
		if (inBody && qualified && tokens[i - 3]?.rawText !== '.' && tokens[i - 3]?.rawText !== '!') {
			const receiver = tokenName(tokens[i - 2])?.toLowerCase();
			const type = receiver ? normalizeType((types ??= typeEnvironmentFor(symbols, member)).get(receiver)) : undefined;
			if (type && type !== 'object' && type !== 'variant' && type.split('.').at(-1) !== moduleName.toLowerCase()) { continue; }
		}
		references.add(name);
	}
	return references;
}

function shiftDiagnostic(d: VbaDiagnostic, shift: number, newMemberStart: number): VbaDiagnostic {
	if (shift === 0) {
		return d;
	}
	return {
		...d,
		span: { start: d.span.start + shift, end: d.span.end + shift },
		walkMemberStart: newMemberStart,
	};
}

export function analyzeModuleRulesIncremental(
	source: string,
	opts: AnalyzeModuleOptions,
	prev: ModuleRulesIncrementalState | undefined,
	externalFingerprint: readonly unknown[],
): ModuleRulesIncrementalResult {
	checkAnalysisCancellation(opts);
	const fingerprint = [opts.errorsOnly === true, ...externalFingerprint];
	const mod = opts.parsedModule ?? parseModule(source);
	const members = procedureMembers(mod);
	const symbols = buildModuleSymbols(opts.moduleName ?? 'Module', opts.moduleKind ?? 'standard', source, { parsedModule: mod, conditionalCompilation: opts.conditionalCompilation });
	const moduleShadows = new Set<string>([...(symbols.root.children ?? []), ...(opts.projectVisibleSymbols ?? [])].map(s => s.name.toLowerCase()).filter(name => name === 'debug' || name === 'msgbox'));
	for (const name of opts.knownProcedures ?? []) { if (/^(debug|msgbox)$/i.test(name)) { moduleShadows.add(name.toLowerCase()); } }
	const records = members.map(member => {
		checkAnalysisCancellation(opts);
		const shadows = new Set(moduleShadows);
		for (const symbol of procedureSymbolFor(symbols, member)?.children ?? []) { if (/^(debug|msgbox)$/i.test(symbol.name)) { shadows.add(symbol.name.toLowerCase()); } }
		return recordFor(member, source, shadows);
	});
	const envelope = envelopeOf(source, records);
	if (prev && prev.envelope === envelope && sameFingerprint(prev.fingerprint, fingerprint)) {
		for (let i = 0; i < records.length; i++) {
			if (records[i].bodyText === prev.procedures[i]?.bodyText) {
				records[i].references = prev.procedures[i].references;
			}
		}
	}
	const moduleVariables = new Set((symbols.root.children ?? [])
		.filter(s => s.kind === 'moduleVariable').map(s => s.name.toLowerCase()));
	const writtenNames = new Set([...writtenNamesIn(source)].filter(name => moduleVariables.has(name)));

	const maxDirtyProcedures = Math.max(MIN_DIRTY_PROCEDURES, Math.floor(members.length / 2));
	const surface = prev && sameFingerprint(prev.fingerprint, fingerprint)
		? prev.envelope === envelope && prev.procedures.length === records.length
			? { previousIndices: records.map((_, i) => i), changedHeaders: new Set<number>(), changedClosers: new Set<number>(), changedNames: new Set<string>() }
			: matchDeclarationSurface(source, mod, records, prev)
		: undefined;
	let dirty: Set<number> | undefined;
	if (
		prev &&
		sameFingerprint(prev.fingerprint, fingerprint) &&
		surface &&
		!records.some((record, i) => record.bodyText !== prev.procedures[surface.previousIndices[i]]?.bodyText &&
			/^[ \t]*#/m.test(record.bodyText + '\n' + (prev.procedures[surface.previousIndices[i]]?.bodyText ?? ''))) &&
		prev.writtenNames.size === writtenNames.size && [...writtenNames].every(name => prev.writtenNames.has(name))
	) {
		dirty = new Set<number>();
		for (let i = 0; i < records.length; i += 1) {
			const previousIndex = surface.previousIndices[i];
			if (surface.changedHeaders.has(i) || surface.changedClosers.has(i) || records[i].bodyText !== prev.procedures[previousIndex]?.bodyText) {
				dirty.add(i);
				continue;
			}
			// Cached diagnostics are only spliceable when a uniform per-procedure
			// offset shift is valid: every span inside the procedure, and no
			// structured `data` payload (which may embed offsets elsewhere in the
			// module). Otherwise re-walk the procedure.
			const cached = prev.walkDiagnosticsByProcedure[previousIndex] ?? [];
			const old = prev.procedures[previousIndex];
			const unspliceable = cached.some(
				(d) => d.data !== undefined || d.span.start < old.start || d.span.end > old.end,
			);
			if (unspliceable) {
				dirty.add(i);
			}
		}
		// A changed callee can alter known return values and ByRef effects in
		// callers. References respect local/type binding; unknown receivers and
		// string-based call names stay conservative.
		const effectsChanged = new Set([...dirty].filter(i => surface.changedHeaders.has(i) || records[i].effectText !== prev.procedures[surface.previousIndices[i]]?.effectText));
		// Return literals and replayed member operations affect direct callers.
		// Only argument preservation and module-variable effects can change the
		// facts a caller exports to its own callers. Compare the same facts the
		// rules consume, rather than flooding a whole call tree for every edit.
		const transitive = new Set(surface.changedNames);
		const oldKeeps = calleeKeepsArgument(prev.source);
		const newKeeps = calleeKeepsArgument(source);
		for (const i of dirty) {
			checkAnalysisCancellation(opts);
			if (!surface.changedHeaders.has(i) && records[i].bodyText === prev.procedures[surface.previousIndices[i]]?.bodyText) { continue; }
			const member = members[i];
			if (member.params.some((_, index) => oldKeeps(member.name, index) !== newKeeps(member.name, index))
				|| calleeModuleVariableEffects(prev.source, prev.symbols, member.name) !== calleeModuleVariableEffects(source, symbols, member.name)) {
				transitive.add(member.name.toLowerCase());
				effectsChanged.add(i);
			}
			if (!effectsChanged.has(i) && member.params.some((_, index) =>
				calleeMemberCallFingerprint(prev.source, member.name, index) !== calleeMemberCallFingerprint(source, member.name, index))) { effectsChanged.add(i); }
		}
		// Skip dependencies only when all compared caller facts stayed the same.
		// Other edits reuse name sets for unchanged bodies.
		const references = (i: number): ReadonlySet<string> => records[i].references ??= procedureReferences(source, members[i], symbols, opts.moduleName ?? 'Module');
		if (effectsChanged.size > 0 || surface.changedNames.size > 0) {
			// Traverse the reverse graph once. Repeatedly comparing every body
			// with every growing callee set made fan-out itself a long pause.
			const callableNames = new Set([...members.map(member => member.name.toLowerCase()), ...surface.changedNames]);
			const callers = new Map<string, number[]>();
			for (let i = 0; i < members.length; i++) {
				checkAnalysisCancellation(opts);
				for (const name of references(i)) {
					if (!callableNames.has(name)) { continue; }
					const bucket = callers.get(name);
					if (bucket) { bucket.push(i); } else { callers.set(name, [i]); }
				}
			}
			const pending = [...effectsChanged].map(i => ({ name: members[i].name.toLowerCase(), propagate: transitive.has(members[i].name.toLowerCase()) }));
			for (const name of surface.changedNames) { pending.push({ name, propagate: true }); }
			const visited = new Set<string>();
			for (let next = 0; next < pending.length && dirty.size <= maxDirtyProcedures; next++) {
				const { name, propagate } = pending[next];
				const key = `${name}:${propagate}`;
				if (visited.has(key)) { continue; }
				visited.add(key);
				for (const caller of callers.get(name) ?? []) {
					dirty.add(caller);
					if (propagate) { pending.push({ name: members[caller].name.toLowerCase(), propagate: true }); }
				}
			}
		}
		if (dirty && dirty.size > maxDirtyProcedures) {
			dirty = undefined;
		}
	}

	if (dirty === undefined) {
		const diagnostics = analyzeModule(source, { ...opts, parsedModule: mod, walkProcedureFilter: undefined });
		return {
			diagnostics,
			state: stateFrom(source, symbols, envelope, writtenNames, records, members, diagnostics, fingerprint),
			mode: 'full',
		};
	}

	const dirtyStarts = new Set<number>([...dirty].map((i) => members[i].span.start));
	const fresh = analyzeModule(source, {
		...opts,
		parsedModule: mod,
		walkProcedureFilter: (member) => dirtyStarts.has(member.span.start),
	});

	const walkByStart = new Map<number, VbaDiagnostic[]>(
		members.map((member) => [member.span.start, []]),
	);
	const diagnostics: VbaDiagnostic[] = [];
	for (const d of fresh) {
		diagnostics.push(d);
		if (d.origin === 'walk' && d.walkMemberStart !== undefined) {
			walkByStart.get(d.walkMemberStart)?.push(d);
		}
	}
	for (let i = 0; i < records.length; i += 1) {
		if (dirty.has(i)) {
			continue;
		}
		const previousIndex = surface!.previousIndices[i];
		const shift = records[i].start - prev!.procedures[previousIndex].start;
		const bucket = walkByStart.get(members[i].span.start);
		for (const d of prev!.walkDiagnosticsByProcedure[previousIndex] ?? []) {
			const shifted = shiftDiagnostic(d, shift, records[i].start);
			diagnostics.push(shifted);
			bucket?.push(shifted);
		}
	}

	return {
		diagnostics,
		state: {
			source,
			symbols,
			envelope,
			writtenNames,
			procedures: records,
			walkDiagnosticsByProcedure: members.map((member) => walkByStart.get(member.span.start) ?? []),
			fingerprint,
		},
		mode: 'incremental',
	};
}

function stateFrom(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	envelope: string,
	writtenNames: ReadonlySet<string>,
	records: ProcedureRecord[],
	members: readonly ProcedureNode[],
	diagnostics: readonly VbaDiagnostic[],
	fingerprint: readonly unknown[],
): ModuleRulesIncrementalState {
	const walkByStart = new Map<number, VbaDiagnostic[]>(
		members.map((member) => [member.span.start, []]),
	);
	for (const d of diagnostics) {
		if (d.origin === 'walk' && d.walkMemberStart !== undefined) {
			walkByStart.get(d.walkMemberStart)?.push(d);
		}
	}
	return {
		source,
		symbols,
		envelope,
		writtenNames,
		procedures: records,
		walkDiagnosticsByProcedure: members.map((member) => walkByStart.get(member.span.start) ?? []),
		fingerprint,
	};
}
