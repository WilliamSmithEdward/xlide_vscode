// Semantic tokens and lookup for type names.
//
// TextMate grammar can color only static tokens. This pure resolver marks parsed
// declaration type positions and `New` expressions after resolving the type
// name against project, VBA primitive, and host object-model type metadata.

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
import {
	resolveTypeName,
	type TypeCompletion,
	type TypeCompletionContext,
	type TypeCompletionKind,
} from '../completion/typeCompletion';

export type TypeSemanticTokenType = 'class' | 'enum' | 'struct' | 'type';

export interface TypeSemanticToken {
	name: string;
	tokenType: TypeSemanticTokenType;
	span: Span;
}

export interface ResolvedTypeReference extends TypeCompletion {
	span: Span;
}

type TypeNameHit = { name: string; span: Span };

function tokenTypeForCompletionKind(kind: TypeCompletionKind): TypeSemanticTokenType {
	switch (kind) {
		case 'class':
		case 'document':
		case 'userform':
		case 'host':
			return 'class';
		case 'enum':
			return 'enum';
		case 'userType':
			return 'struct';
		case 'primitive':
		case 'ambiguous':
			return 'type';
	}
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

function collectVariableGroup(
	source: string,
	group: VariableGroupNode,
	out: TypeNameHit[],
): void {
	for (const decl of group.declarations) {
		collectVariableDecl(source, decl, out);
	}
}

function collectVariableDecl(
	source: string,
	decl: VariableDeclNode,
	out: TypeNameHit[],
): void {
	if (!decl.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, decl.span));
}

function collectTypeField(
	source: string,
	field: TypeFieldNode,
	out: TypeNameHit[],
): void {
	if (!field.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, field.span));
}

function collectParameter(
	source: string,
	param: ParameterNode,
	out: TypeNameHit[],
): void {
	if (!param.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, param.span));
}

function collectBody(
	source: string,
	body: BodyNode[],
	out: TypeNameHit[],
): void {
	for (const node of body) {
		if (node.kind === 'VariableGroup') {
			collectVariableGroup(source, node, out);
		} else if (node.kind === 'Statement') {
			collectStatement(source, node, out);
		} else if ('body' in node && Array.isArray(node.body)) {
			collectBody(source, node.body, out);
		}
	}
}

function collectStatement(
	source: string,
	stmt: StatementNode,
	out: TypeNameHit[],
): void {
	for (const hit of typeNameSpansAfterNew(source, stmt.span)) {
		out.push(hit);
	}
}

function collectProcedure(
	source: string,
	proc: ProcedureNode,
	out: TypeNameHit[],
): void {
	for (const param of proc.params) {
		collectParameter(source, param, out);
	}
	if (proc.returnType) {
		pushTypeHit(out, returnTypeNameSpan(source, proc));
	}
	collectBody(source, proc.body, out);
}

function collectModule(
	source: string,
	mod: ModuleNode,
): TypeNameHit[] {
	const out: TypeNameHit[] = [];
	for (const member of mod.members) {
		if (member.kind === 'VariableGroup') {
			collectVariableGroup(source, member, out);
		} else if (member.kind === 'Type') {
			for (const field of member.fields) {
				collectTypeField(source, field, out);
			}
		} else if (member.kind === 'Procedure') {
			collectProcedure(source, member, out);
		}
	}
	return out.sort((a, b) => a.span.start - b.span.start || a.span.end - b.span.end);
}

function pushTypeHit(out: TypeNameHit[], hit: TypeNameHit | undefined): void {
	if (hit) {
		out.push(hit);
	}
}

function semanticTokenForHit(
	ctx: TypeCompletionContext,
	hit: TypeNameHit,
): TypeSemanticToken | undefined {
	const resolved = resolveTypeName(hit.name, ctx);
	if (!resolved) {
		return undefined;
	}
	return {
		name: hit.name,
		tokenType: tokenTypeForCompletionKind(resolved.kind),
		span: hit.span,
	};
}

function collectTypeNameHits(source: string): TypeNameHit[] {
	return collectModule(source, parseModule(source));
}

export function resolveTypeSemanticTokens(
	source: string,
	ctx: TypeCompletionContext = {},
): TypeSemanticToken[] {
	return collectTypeNameHits(source)
		.map((hit) => semanticTokenForHit(ctx, hit))
		.filter((token): token is TypeSemanticToken => Boolean(token));
}

export function resolveTypeReferenceAt(
	source: string,
	offset: number,
	ctx: TypeCompletionContext = {},
): ResolvedTypeReference | undefined {
	const hit = collectTypeNameHits(source).find(
		(candidate) => offset >= candidate.span.start && offset <= candidate.span.end,
	);
	if (!hit) {
		return undefined;
	}
	const resolved = resolveTypeName(hit.name, ctx);
	if (!resolved) {
		return undefined;
	}
	return { ...resolved, span: hit.span };
}
