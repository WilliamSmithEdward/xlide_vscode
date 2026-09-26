// Rule: conditional-compilation directive forms the VBE refuses (issue #130).
// Measured in Excel 16.0 (build 20326, 2026-09-26):
//
//  - duplicate-const-directive: `#Const FEATURE = 1` twice in one module ->
//    "Duplicate definition".
//  - directive-trailing-statement: code after a colon on a directive line,
//    `#If VBA7 Then: Debug.Print 1` -> "An # ElseIf, # Else, or # EndIf must
//    be preceded by an # If clause" (the colon ends the directive, and what
//    follows is no longer part of it).

import type { ConditionalDirectiveNode, ModuleNode, BodyNode } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { tokenizeCached } from '../../lexer/tokenize';

export function checkDirectiveForms(source: string, mod: ModuleNode, push: PushFn): void {
	const directives: ConditionalDirectiveNode[] = [];
	const visit = (nodes: readonly BodyNode[]): void => {
		for (const node of nodes) {
			if (node.kind === 'ConditionalDirective') {
				directives.push(node);
			} else if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
		}
	};
	for (const member of mod.members) {
		if (member.kind === 'ConditionalDirective') {
			directives.push(member);
		} else if (member.kind === 'Procedure') {
			visit(member.body);
		} else if (member.kind === 'Enum' || member.kind === 'Type') {
			for (const directive of member.directives ?? []) {
				directives.push(directive);
			}
		}
	}
	const defined = new Map<string, ConditionalDirectiveNode>();
	for (const directive of directives) {
		if (directive.directiveKind === 'Const' && directive.name) {
			const lower = directive.name.toLowerCase();
			const earlier = defined.get(lower);
			if (earlier) {
				push(
					'duplicateConstDirective',
					`'#Const ${directive.name}' is already defined in this module. This is a VBE compile error: Duplicate definition.`,
					directive.nameSpan ?? directive.span,
				);
			} else {
				defined.set(lower, directive);
			}
		}
	}
	// A directive line that a colon continues: the tokens after the colon on
	// the directive's own physical line.
	const tokens = tokenizeCached(source);
	for (const directive of directives) {
		const lineEnd = lineEndAfter(source, directive.span.end);
		let colon = -1;
		for (let i = 0; i < tokens.length; i++) {
			const tok = tokens[i];
			if (tok.start < directive.span.end) {
				continue;
			}
			if (tok.start >= lineEnd) {
				break;
			}
			if (tok.kind === 'colon') {
				colon = i;
				break;
			}
		}
		if (colon < 0) {
			continue;
		}
		const next = tokens[colon + 1];
		if (next && next.kind !== 'newline' && next.kind !== 'comment' && next.start < lineEnd) {
			push(
				'directiveTrailingStatement',
				`A compiler directive takes the whole line: nothing may follow the ':' after it. This is a VBE compile error (the VBE reports "An # ElseIf, # Else, or # EndIf must be preceded by an # If clause").`,
				{ start: next.start, end: lineEnd },
			);
		}
	}
}

function lineEndAfter(source: string, from: number): number {
	for (let i = from; i < source.length; i++) {
		if (source[i] === '\r' || source[i] === '\n') {
			return i;
		}
	}
	return source.length;
}
