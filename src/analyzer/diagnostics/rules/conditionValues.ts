// Rule family: a Collection, an array or a Variant holding an array read as a
// condition or a Boolean operand (issue #424, each measured in Excel 16.0).
//
// A Collection's default member Item needs an index. As the condition of an
// If, a loop or IIf it raises 450 when it runs; as a Select Case subject or
// an operand of Not, And or Or it does not compile ("Argument not optional").
// An array declared as one is a compile "Type mismatch" as a condition or a
// Select Case subject, and IIf raises 13 on it; `Not a` runs (it is the
// `Not Not a` idiom), and And or Or on it is non-scalar-binary-operand's. A
// Variant holding an array raises 13 in every one of these places.
//
// The forms are read on each statement, a block If's own line and each
// ElseIf line, a loop's or a Select's opening line and a Do's Loop line.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { BodyNode, LeafStatementNode, ModuleNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { blockHeaderStatements } from '../blockHeaders';
import { heldObjectsAt } from '../heldObjects';
import { conditionOperands, type ConditionForm } from '../conditionOperands';
import { normalizeType } from '../typeInference';
import { activeModuleMembers, isInactiveNode, statementTokensAfterLeadingLabel } from '../walker';
import { knownArrayShapesAt, moduleOptionBase } from './arrays';

const WHERE: Readonly<Record<ConditionForm, string>> = {
	condition: 'the condition',
	select: 'Select Case',
	iif: 'IIf',
	not: "'Not'",
	logical: 'the Boolean operator',
};

export function checkConditionValues(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const optionBase = moduleOptionBase(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = (procedureSymbolFor(symbols, member)?.children ?? []).filter((child) => child.kind === 'localVariable');
		const collections = new Map(locals.filter((child) => !child.isArray && normalizeType(child.asType) === 'collection').map((child) => [child.name.toLowerCase(), child.isAutoInstantiated === true]));
		const arrays = new Set(locals.filter((child) => child.isArray).map((child) => child.name.toLowerCase()));
		const variants = new Set(locals.filter((child) => !child.isArray && (normalizeType(child.asType) ?? 'variant') === 'variant').map((child) => child.name.toLowerCase()));
		if (collections.size === 0 && arrays.size === 0 && variants.size === 0) {
			continue;
		}
		let shapesAt: ReturnType<typeof knownArrayShapesAt> | undefined;
		// A Collection set here holds one, as an `As New` one always does
		// (issue #415, measured in Excel 16.0).
		let heldAt: ReturnType<typeof heldObjectsAt> | undefined;
		const holdsOne = (stmt: LeafStatementNode, lower: string): boolean =>
			collections.get(lower) === true || (heldAt ??= heldObjectsAt(source, member, symbols, activity))(stmt).classes.get(lower)?.toLowerCase() === 'collection';
		const check = (stmt: LeafStatementNode): void => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span).filter((tok) => tok.kind !== 'comment');
			for (const { index, form } of conditionOperands(toks)) {
				const tok = toks[index];
				const lower = tok.rawText.toLowerCase();
				const at = { start: stmt.span.start + tok.start, end: stmt.span.start + tok.end };
				const where = WHERE[form];
				if (collections.has(lower)) {
					if (form === 'select' || form === 'not' || form === 'logical') {
						push('collectionOperand', `'${tok.rawText}' is a Collection: its default member Item needs an index, so ${where} has no value to work on. This is a VBE compile error: Argument not optional.`, at);
					} else if (holdsOne(stmt, lower)) {
						push('objectDefaultValue', `'${tok.rawText}' is a Collection: its default member Item needs an index, so ${where} has no value to read. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`, at);
					}
				} else if (arrays.has(lower)) {
					if (form === 'condition' || form === 'select') {
						push('nonScalarBinaryOperand', `'${tok.rawText}' is declared as an array, which ${where} cannot read as one value. This will fail to compile with 'Type mismatch'.`, at);
					} else if (form === 'iif') {
						push('variantValueMisuse', `'${tok.rawText}' is an array, which IIf cannot read as its condition. This will raise Run-time error '13': Type mismatch.`, at);
					}
				} else if (variants.has(lower)) {
					const shape = (shapesAt ??= knownArrayShapesAt(source, symbols, member, activity, optionBase))(stmt).get(lower);
					if (shape) {
						push('variantValueMisuse', `'${tok.rawText}' holds an array from ${shape.origin} here, which ${where} cannot read as one value. This will raise Run-time error '13': Type mismatch.`, at);
					}
				}
			}
		};
		const visit = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (isInactiveNode(activity, node)) {
					continue;
				}
				if (isLeafStatement(node)) {
					check(node);
					continue;
				}
				if (!('body' in node) || !Array.isArray(node.body)) {
					continue;
				}
				const { before, after } = blockHeaderStatements(source, node);
				const opening = node.kind === 'IfBlock' ? node.branches[0]?.headerSpan : undefined;
				const header = before ?? (opening ? { kind: 'Statement' as const, span: opening, raw: source.slice(opening.start, opening.end) } : undefined);
				if (header) {
					check(header);
				}
				visit(node.body as BodyNode[]);
				if (after) {
					check(after);
				}
			}
		};
		visit(member.body);
	}
}
