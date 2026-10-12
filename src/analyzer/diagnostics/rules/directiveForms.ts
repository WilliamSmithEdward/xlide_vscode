// Rule: conditional-compilation directive forms the VBE refuses (issue #130).
// Measured in Excel 16.0 (build 20326, 2026-09-26):
//
//  - duplicate-const-directive: `#Const FEATURE = 1` twice in one module ->
//    "Duplicate definition".
//  - directive-trailing-statement: code after a colon on a directive line,
//    `#If VBA7 Then: Debug.Print 1` -> "An # ElseIf, # Else, or # EndIf must
//    be preceded by an # If clause" (the colon ends the directive, and what
//    follows is no longer part of it).
//  - null-directive-condition: `#If Null Then`, `#If Null = 1 Then`, or `#If N`
//    after `#Const N = Null` -> "Invalid use of Null" (issue #208).

import type { ConditionalDirectiveNode, ModuleNode, BodyNode } from '../../parser/nodes';
import {
	nullConditionDirectives,
	type ConditionalActivityTracker,
	type ConditionalCompilationEnvironment,
} from '../../conditional/conditionalCompilation';
import type { PushFn } from '../analysisContext';
import { tokenizeCached } from '../../lexer/tokenize';
import { colonDirectiveStarts, directiveColon, firstTokenAtOrAfter, tokenWord as tokenText } from '../../lexer/tokenHelpers';
import { lineEndAtOrAfter } from '../../../vbaSourceScan';

/** What a directive line a colon follows fails to do, by its keyword (issue #1323). */
const COLON_DIRECTIVE_EFFECT: Readonly<Record<string, string>> = {
	if: 'opens no #If block, so the #Else, #ElseIf or #End If after it has no #If',
	elseif: 'starts no new arm: it belongs to the arm above it',
	else: 'starts no new arm: it belongs to the arm above it',
	end: 'closes no #If block',
	endif: 'closes no #If block',
	const: 'defines no constant',
};

export function checkDirectiveForms(
	source: string,
	mod: ModuleNode,
	conditionalCompilation: ConditionalCompilationEnvironment | undefined,
	push: PushFn,
	activity?: ConditionalActivityTracker,
): void {
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
	// A directive line a statement colon follows is no directive: the VBE
	// reads the whole line as code, refused where it is compiled and skipped
	// with a skipped arm (issue #1323). The parser and analyzeVbaStructure
	// read the same lines, so the #If pairing a refused line breaks is theirs.
	const starts = colonDirectiveStarts(source);
	if (starts.size > 0) {
		const tokens = tokenizeCached(source);
		for (const start of starts) {
			if (activity?.isInactive({ start, end: start })) {
				continue;
			}
			const hash = firstTokenAtOrAfter(tokens, start);
			const colon = directiveColon(tokens, hash);
			if (!colon) {
				continue;
			}
			const keyword = tokenText(tokens[hash + 1]);
			const effect = COLON_DIRECTIVE_EFFECT[keyword] ?? COLON_DIRECTIVE_EFFECT.if;
			// What follows the colon, or the colon itself when nothing does.
			const lineEnd = lineEndAtOrAfter(source, colon.start);
			const after = tokens[tokens.indexOf(colon, hash) + 1];
			const from = after && after.kind !== 'newline' && after.start < lineEnd ? after.start : colon.start;
			push(
				'directiveTrailingStatement',
				`A ':' after a compiler directive makes the whole line no directive, so it ${effect}. A directive takes its line alone: drop the ':', and start any comment with '. This is a VBE compile error.`,
				{ start: from, end: from === colon.start ? colon.end : lineEnd },
			);
		}
	}
	for (const directive of nullConditionDirectives(mod, conditionalCompilation)) {
		push(
			'nullDirectiveCondition',
			`This #${directive.directiveKind} condition is Null, which is neither True nor False. This is a VBE compile error: Invalid use of Null.`,
			directive.span,
		);
	}
}
