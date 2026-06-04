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
	qualifier?: string;
	qualifierSpan?: Span;
	fullSpan?: Span;
}

export type TypeNameReferenceKind =
	| 'declaration'
	| 'newDeclaration'
	| 'newExpression'
	| 'typeOfIs'
	| 'implements';

export interface TypeNameReference {
	name: string;
	qualifier?: string;
	qualifierSpan?: Span;
	span: Span;
	fullSpan?: Span;
	kind: TypeNameReferenceKind;
}

function tokenTypeForCompletionKind(kind: TypeCompletionKind): TypeSemanticTokenType {
	switch (kind) {
		case 'class':
		case 'document':
		case 'userform':
		case 'host':
		case 'external':
			return 'class';
		case 'enum':
			return 'enum';
		case 'userType':
			return 'struct';
		case 'primitive':
		case 'ambiguous':
		case 'module':
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

function typeNameReferenceFromTokens(
	toks: readonly VbaToken[],
	typeIndex: number,
	base: number,
	kind: TypeNameReferenceKind,
): TypeNameReference | undefined {
	const firstName = tokenName(toks[typeIndex]);
	if (!firstName) {
		return undefined;
	}
	const firstSpan = {
		start: base + toks[typeIndex].start,
		end: base + toks[typeIndex].end,
	};
	const memberName = toks[typeIndex + 1]?.rawText === '.'
		? tokenName(toks[typeIndex + 2])
		: undefined;
	if (!memberName) {
		return {
			name: firstName,
			span: firstSpan,
			kind,
		};
	}
	const memberSpan = {
		start: base + toks[typeIndex + 2].start,
		end: base + toks[typeIndex + 2].end,
	};
	return {
		name: memberName,
		qualifier: firstName,
		qualifierSpan: firstSpan,
		span: memberSpan,
		fullSpan: {
			start: firstSpan.start,
			end: memberSpan.end,
		},
		kind,
	};
}

function typeNameSpanAfterAs(source: string, span: Span): TypeNameReference | undefined {
	const toks = codeTokens(source, span);
	for (let i = 0; i < toks.length; i++) {
		if ((toks[i].canonicalText ?? toks[i].rawText).toLowerCase() !== 'as') {
			continue;
		}
		let typeIndex = i + 1;
		let kind: TypeNameReferenceKind = 'declaration';
		if ((toks[typeIndex]?.canonicalText ?? toks[typeIndex]?.rawText ?? '').toLowerCase() === 'new') {
			typeIndex++;
			kind = 'newDeclaration';
		}
		const ref = typeNameReferenceFromTokens(toks, typeIndex, span.start, kind);
		if (ref) {
			return ref;
		}
	}
	return undefined;
}

function typeNameSpansAfterNew(source: string, span: Span): TypeNameReference[] {
	const toks = codeTokens(source, span);
	const out: TypeNameReference[] = [];
	for (let i = 0; i < toks.length; i++) {
		if ((toks[i].canonicalText ?? toks[i].rawText).toLowerCase() !== 'new') {
			continue;
		}
		const ref = typeNameReferenceFromTokens(toks, i + 1, span.start, 'newExpression');
		if (ref) {
			out.push(ref);
		}
	}
	return out;
}

function typeNameSpansAfterTypeOfIs(source: string, span: Span): TypeNameReference[] {
	const toks = codeTokens(source, span);
	const out: TypeNameReference[] = [];
	let sawTypeOf = false;
	for (let i = 0; i < toks.length; i++) {
		const lower = (toks[i].canonicalText ?? toks[i].rawText).toLowerCase();
		if (lower === 'typeof') {
			sawTypeOf = true;
			continue;
		}
		if (!sawTypeOf || lower !== 'is') {
			continue;
		}
		const ref = typeNameReferenceFromTokens(toks, i + 1, span.start, 'typeOfIs');
		if (ref) {
			out.push(ref);
		}
		sawTypeOf = false;
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

function returnTypeNameSpan(source: string, proc: ProcedureNode): TypeNameReference | undefined {
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
		const ref = typeNameReferenceFromTokens(toks, i + 1, header.start, 'declaration');
		if (!ref) {
			return undefined;
		}
		return ref;
	}
	return undefined;
}

function collectVariableGroup(
	source: string,
	group: VariableGroupNode,
	out: TypeNameReference[],
): void {
	for (const decl of group.declarations) {
		collectVariableDecl(source, decl, out);
	}
}

function collectVariableDecl(
	source: string,
	decl: VariableDeclNode,
	out: TypeNameReference[],
): void {
	if (!decl.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, decl.span));
}

function collectTypeField(
	source: string,
	field: TypeFieldNode,
	out: TypeNameReference[],
): void {
	if (!field.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, field.span));
}

function collectParameter(
	source: string,
	param: ParameterNode,
	out: TypeNameReference[],
): void {
	if (!param.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, param.span));
}

function collectBody(
	source: string,
	body: BodyNode[],
	out: TypeNameReference[],
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
	out: TypeNameReference[],
): void {
	for (const hit of typeNameSpansAfterNew(source, stmt.span)) {
		out.push(hit);
	}
	for (const hit of typeNameSpansAfterTypeOfIs(source, stmt.span)) {
		out.push(hit);
	}
}

function collectImplements(source: string, out: TypeNameReference[]): void {
	let lineStart = 0;
	while (lineStart <= source.length) {
		let lineEnd = source.indexOf('\n', lineStart);
		if (lineEnd < 0) {
			lineEnd = source.length;
		}
		let line = source.slice(lineStart, lineEnd);
		if (line.endsWith('\r')) {
			line = line.slice(0, -1);
		}
		const code = line.replace(/'.*$/, '');
		const match = /^\s*Implements\s+([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)\b/i.exec(code);
		if (match) {
			const rawName = match[1];
			const column = line.indexOf(rawName, match.index);
			const dot = rawName.indexOf('.');
			if (column >= 0 && dot > 0) {
				out.push({
					name: rawName.slice(dot + 1),
					qualifier: rawName.slice(0, dot),
					qualifierSpan: {
						start: lineStart + column,
						end: lineStart + column + dot,
					},
					span: {
						start: lineStart + column + dot + 1,
						end: lineStart + column + rawName.length,
					},
					fullSpan: {
						start: lineStart + column,
						end: lineStart + column + rawName.length,
					},
					kind: 'implements',
				});
			} else if (column >= 0) {
				out.push({
					name: rawName,
					span: {
						start: lineStart + column,
						end: lineStart + column + rawName.length,
					},
					kind: 'implements',
				});
			}
		}
		if (lineEnd === source.length) {
			break;
		}
		lineStart = lineEnd + 1;
	}
}

function collectProcedure(
	source: string,
	proc: ProcedureNode,
	out: TypeNameReference[],
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
): TypeNameReference[] {
	const out: TypeNameReference[] = [];
	collectImplements(source, out);
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

function pushTypeHit(
	out: TypeNameReference[],
	hit: TypeNameReference | undefined,
): void {
	if (hit) {
		out.push(hit);
	}
}

function semanticTokenForHit(
	ctx: TypeCompletionContext,
	hit: TypeNameReference,
): TypeSemanticToken | undefined {
	const resolved = resolveTypeName(typeReferenceLookupName(hit), ctx);
	if (!resolved) {
		return undefined;
	}
	return {
		name: hit.name,
		tokenType: tokenTypeForCompletionKind(resolved.kind),
		span: hit.span,
	};
}

export function typeReferenceLookupName(hit: TypeNameReference): string {
	return hit.qualifier ? `${hit.qualifier}.${hit.name}` : hit.name;
}

export function collectTypeNameReferences(source: string): TypeNameReference[] {
	return collectModule(source, parseModule(source));
}

export function resolveTypeSemanticTokens(
	source: string,
	ctx: TypeCompletionContext = {},
): TypeSemanticToken[] {
	return collectTypeNameReferences(source)
		.map((hit) => semanticTokenForHit(ctx, hit))
		.filter((token): token is TypeSemanticToken => Boolean(token));
}

export function resolveTypeReferenceAt(
	source: string,
	offset: number,
	ctx: TypeCompletionContext = {},
): ResolvedTypeReference | undefined {
	const hit = collectTypeNameReferences(source).find(
		(candidate) => offset >= candidate.span.start && offset <= candidate.span.end,
	);
	if (!hit) {
		return undefined;
	}
	const resolved = resolveTypeName(typeReferenceLookupName(hit), ctx);
	if (!resolved) {
		return undefined;
	}
	return {
		...resolved,
		span: hit.span,
		qualifier: hit.qualifier,
		qualifierSpan: hit.qualifierSpan,
		fullSpan: hit.fullSpan,
	};
}
