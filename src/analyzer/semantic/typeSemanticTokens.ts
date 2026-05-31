// Semantic tokens for project-defined type names.
//
// TextMate grammar can color only static tokens. Workbook classes, UserForms,
// document modules, UDTs, and enums are dynamic project symbols, so this pure
// resolver marks only parsed declaration type positions whose type name resolves
// through the project binder.

import type {
	BodyNode,
	ModuleNode,
	ParameterNode,
	ProcedureNode,
	Span,
	TypeFieldNode,
	VariableDeclNode,
	VariableGroupNode,
	StatementNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import type { VbaProjectTypeName } from '../symbols/symbolModel';

export type ProjectTypeSemanticTokenType = 'class' | 'enum' | 'struct' | 'type';

export interface ProjectTypeSemanticToken {
	name: string;
	tokenType: ProjectTypeSemanticTokenType;
	span: Span;
}

type TypeCandidate = {
	name: string;
	tokenType: ProjectTypeSemanticTokenType;
};

function tokenTypeForProjectType(kind: VbaProjectTypeName['kind']): ProjectTypeSemanticTokenType {
	switch (kind) {
		case 'class':
		case 'document':
		case 'userform':
			return 'class';
		case 'enum':
			return 'enum';
		case 'userType':
			return 'struct';
	}
}

function typeCandidatesByName(
	projectTypes: readonly VbaProjectTypeName[],
): ReadonlyMap<string, TypeCandidate> {
	const grouped = new Map<string, Set<ProjectTypeSemanticTokenType>>();
	const casing = new Map<string, string>();
	for (const t of projectTypes) {
		const key = t.name.toLowerCase();
		casing.set(key, casing.get(key) ?? t.name);
		const set = grouped.get(key) ?? new Set<ProjectTypeSemanticTokenType>();
		set.add(tokenTypeForProjectType(t.kind));
		grouped.set(key, set);
	}

	const out = new Map<string, TypeCandidate>();
	for (const [key, types] of grouped) {
		out.set(key, {
			name: casing.get(key) ?? key,
			tokenType: types.size === 1 ? [...types][0] : 'type',
		});
	}
	return out;
}

function codeTokens(source: string, span: Span): VbaToken[] {
	return tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
}

function tokenName(tok: VbaToken | undefined): string | undefined {
	if (!tok) {
		return undefined;
	}
	if (tok.kind === 'identifier' || tok.kind === 'keyword') {
		return tok.rawText;
	}
	if (tok.kind === 'bracketedIdentifier') {
		return tok.rawText.slice(1, -1);
	}
	return undefined;
}

function typeNameSpanAfterAs(source: string, span: Span): { name: string; span: Span } | undefined {
	const toks = codeTokens(source, span);
	for (let i = 0; i < toks.length; i++) {
		if ((toks[i].canonicalText ?? toks[i].rawText).toLowerCase() !== 'as') {
			continue;
		}
		let typeIndex = i + 1;
		if ((toks[typeIndex]?.canonicalText ?? toks[typeIndex]?.rawText ?? '').toLowerCase() === 'new') {
			typeIndex++;
		}
		const name = tokenName(toks[typeIndex]);
		if (!name) {
			continue;
		}
		return {
			name,
			span: {
				start: span.start + toks[typeIndex].start,
				end: span.start + toks[typeIndex].end,
			},
		};
	}
	return undefined;
}

function typeNameSpansAfterNew(source: string, span: Span): { name: string; span: Span }[] {
	const toks = codeTokens(source, span);
	const out: { name: string; span: Span }[] = [];
	for (let i = 0; i < toks.length; i++) {
		if ((toks[i].canonicalText ?? toks[i].rawText).toLowerCase() !== 'new') {
			continue;
		}
		const name = tokenName(toks[i + 1]);
		if (!name) {
			continue;
		}
		out.push({
			name,
			span: {
				start: span.start + toks[i + 1].start,
				end: span.start + toks[i + 1].end,
			},
		});
	}
	return out;
}

function headerSpanForProcedure(source: string, proc: ProcedureNode): Span {
	const nl = source.indexOf('\n', proc.span.start);
	return {
		start: proc.span.start,
		end: nl === -1 ? proc.span.end : Math.min(nl, proc.span.end),
	};
}

function returnTypeNameSpan(source: string, proc: ProcedureNode): { name: string; span: Span } | undefined {
	const header = headerSpanForProcedure(source, proc);
	const toks = codeTokens(source, header);
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
			continue;
		}
		if (raw === ')') {
			depth--;
			continue;
		}
		if (depth !== 0 || (toks[i].canonicalText ?? raw).toLowerCase() !== 'as') {
			continue;
		}
		const name = tokenName(toks[i + 1]);
		if (!name) {
			return undefined;
		}
		return {
			name,
			span: {
				start: header.start + toks[i + 1].start,
				end: header.start + toks[i + 1].end,
			},
		};
	}
	return undefined;
}

function pushIfProjectType(
	out: ProjectTypeSemanticToken[],
	candidates: ReadonlyMap<string, TypeCandidate>,
	hit: { name: string; span: Span } | undefined,
): void {
	if (!hit) {
		return;
	}
	const resolved = candidates.get(hit.name.toLowerCase());
	if (!resolved) {
		return;
	}
	out.push({
		name: hit.name,
		tokenType: resolved.tokenType,
		span: hit.span,
	});
}

function collectVariableGroup(
	source: string,
	group: VariableGroupNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	for (const decl of group.declarations) {
		collectVariableDecl(source, decl, candidates, out);
	}
}

function collectVariableDecl(
	source: string,
	decl: VariableDeclNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	if (!decl.asType) {
		return;
	}
	pushIfProjectType(out, candidates, typeNameSpanAfterAs(source, decl.span));
}

function collectTypeField(
	source: string,
	field: TypeFieldNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	if (!field.asType) {
		return;
	}
	pushIfProjectType(out, candidates, typeNameSpanAfterAs(source, field.span));
}

function collectParameter(
	source: string,
	param: ParameterNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	if (!param.asType) {
		return;
	}
	pushIfProjectType(out, candidates, typeNameSpanAfterAs(source, param.span));
}

function collectBody(
	source: string,
	body: BodyNode[],
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	for (const node of body) {
		if (node.kind === 'VariableGroup') {
			collectVariableGroup(source, node, candidates, out);
		} else if (node.kind === 'Statement') {
			collectStatement(source, node, candidates, out);
		} else if ('body' in node && Array.isArray(node.body)) {
			collectBody(source, node.body, candidates, out);
		}
	}
}

function collectStatement(
	source: string,
	stmt: StatementNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	for (const hit of typeNameSpansAfterNew(source, stmt.span)) {
		pushIfProjectType(out, candidates, hit);
	}
}

function collectProcedure(
	source: string,
	proc: ProcedureNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
	out: ProjectTypeSemanticToken[],
): void {
	for (const param of proc.params) {
		collectParameter(source, param, candidates, out);
	}
	if (proc.returnType) {
		pushIfProjectType(out, candidates, returnTypeNameSpan(source, proc));
	}
	collectBody(source, proc.body, candidates, out);
}

function collectModule(
	source: string,
	mod: ModuleNode,
	candidates: ReadonlyMap<string, TypeCandidate>,
): ProjectTypeSemanticToken[] {
	const out: ProjectTypeSemanticToken[] = [];
	for (const member of mod.members) {
		if (member.kind === 'VariableGroup') {
			collectVariableGroup(source, member, candidates, out);
		} else if (member.kind === 'Type') {
			for (const field of member.fields) {
				collectTypeField(source, field, candidates, out);
			}
		} else if (member.kind === 'Procedure') {
			collectProcedure(source, member, candidates, out);
		}
	}
	return out.sort((a, b) => a.span.start - b.span.start || a.span.end - b.span.end);
}

export function resolveProjectTypeSemanticTokens(
	source: string,
	projectTypes: readonly VbaProjectTypeName[],
): ProjectTypeSemanticToken[] {
	if (projectTypes.length === 0) {
		return [];
	}
	const candidates = typeCandidatesByName(projectTypes);
	if (candidates.size === 0) {
		return [];
	}
	return collectModule(source, parseModule(source), candidates);
}
