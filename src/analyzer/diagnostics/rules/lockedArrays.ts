// Rule: an array resized or erased while VBA holds it locked (issue #283,
// each measured in Excel 16.0: error 10, "This array is fixed or temporarily
// locked").
//
// A For Each over an array, a With on one of its elements, and an element
// passed ByRef each lock the array until they end. Inside the loop or the
// With, `Erase a`, `ReDim a(5)`, `ReDim Preserve a(5)` and, for a Variant
// holding the array, `v = Array(9)` raise 10. So does a call that passes
// `a(0)` ByRef beside `a` itself to a procedure of the module that erases or
// ReDims that array parameter. Only a statement the block runs every time
// is judged: one inside an If, a Select or another loop, or after an Exit,
// may not run.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, ProcedureNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { extractCall, isNamedSlot } from '../callExtraction';
import { normalizeType } from '../typeInference';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	forEachStatement,
	isInactiveNode,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { knownArrayShapesAt, moduleOptionBase } from './arrays';

/** Statement heads after which a body's later statements may not run. */
const LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'return', 'end', 'resume', 'on', 'stop', 'error']);

export function checkLockedArrays(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const procedures = new Map<string, ProcedureNode | null>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			const lower = member.name.toLowerCase();
			procedures.set(lower, procedures.has(lower) ? null : member);
		}
	}
	const optionBase = moduleOptionBase(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = (procedureSymbolFor(symbols, member)?.children ?? []).filter((child) => child.kind === 'localVariable' && child.visibility !== 'Static');
		const dynamic = new Set(locals.filter((child) => child.isArray && child.arrayBounds === undefined).map((child) => child.name.toLowerCase()));
		const variants = new Set(locals.filter((child) => !child.isArray && (normalizeType(child.asType) ?? 'variant') === 'variant').map((child) => child.name.toLowerCase()));
		if (dynamic.size === 0 && variants.size === 0) {
			continue;
		}
		let shapesAt: ReturnType<typeof knownArrayShapesAt> | undefined;
		// A Variant locks only while it holds an array with an element to step through.
		const holdsElements = (node: BodyNode, lower: string): boolean => {
			if (dynamic.has(lower)) {
				return true;
			}
			const shape = (shapesAt ??= knownArrayShapesAt(source, symbols, member, activity, optionBase))(node as never).get(lower);
			return shape !== undefined && shape.dims.length > 0 && shape.dims.every((dim) => dim.upper >= dim.lower);
		};
		const visit = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (isInactiveNode(activity, node) || !('body' in node) || !Array.isArray(node.body)) {
					if (isLeafStatement(node) && !isInactiveNode(activity, node)) {
						checkElementPass(node.span, dynamic, procedures, source, activity, push);
					}
					continue;
				}
				if (node.kind === 'ForBlock' && node.each) {
					const lower = node.sourceExpression?.trim().toLowerCase();
					if (lower && /^[a-z_]\w*$/.test(lower) && (dynamic.has(lower) || variants.has(lower)) && holdsElements(node, lower)) {
						reportUnlocks(source, node.body as BodyNode[], lower, variants.has(lower), 'the For Each over it', activity, push);
					}
				}
				if (node.kind === 'WithBlock') {
					const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span));
					const lower = tokenName(header[1])?.toLowerCase();
					if (tokenText(header[0]) === 'with' && lower && dynamic.has(lower) && header[2]?.rawText === '(' && header[header.length - 1]?.rawText === ')') {
						reportUnlocks(source, node.body as BodyNode[], lower, false, 'the With on its element', activity, push);
					}
				}
				visit(node.body as BodyNode[]);
			}
		};
		visit(member.body);
	}
}

/** The statements a block runs every time, up to the first that may leave it. */
function everyTime(source: string, body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined): Array<{ node: BodyNode; toks: VbaToken[] }> {
	const out: Array<{ node: BodyNode; toks: VbaToken[] }> = [];
	for (const node of body) {
		if (isInactiveNode(activity, node) || node.kind === 'ConditionalDirective' || node.kind === 'VariableGroup') {
			continue;
		}
		if (!isLeafStatement(node)) {
			let leaves = false;
			forEachStatement([node], (stmt) => {
				leaves ||= LEAVING_HEADS.has(tokenText(statementTokensAfterLeadingLabel(source, stmt.span)[0]));
			}, activity);
			if (leaves) {
				return out;
			}
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span);
		if (LEAVING_HEADS.has(tokenText(toks[0]))) {
			return out;
		}
		if (!(node.kind === 'Statement' && node.singleLineIfBranches)) {
			out.push({ node, toks });
		}
	}
	return out;
}

/** What in a locked array's block resizes or erases it, each reported. */
function reportUnlocks(
	source: string,
	body: readonly BodyNode[],
	lower: string,
	variant: boolean,
	lock: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const { node, toks } of everyTime(source, body, activity)) {
		const head = tokenText(toks[0]);
		let at: VbaToken | undefined;
		let what = '';
		if (head === 'erase') {
			at = toks.find((tok, i) => i > 0 && tokenName(tok)?.toLowerCase() === lower && toks[i - 1]?.rawText !== '.');
			what = 'Erase cannot free it';
		} else if (head === 'redim') {
			const target = toks[tokenText(toks[1]) === 'preserve' ? 2 : 1];
			at = tokenName(target)?.toLowerCase() === lower ? target : undefined;
			what = 'ReDim cannot resize it';
		} else if (variant && tokenName(toks[0])?.toLowerCase() === lower && toks[1]?.rawText === '=') {
			at = toks[0];
			what = 'an assignment cannot replace it';
		}
		if (at) {
			push('arrayTemporarilyLocked', `'${at.rawText}' is locked by ${lock}, so ${what}. This will raise Run-time error '10': This array is fixed or temporarily locked.`, { start: node.span.start + at.start, end: node.span.start + at.end });
		}
	}
}

/** `Zap a(0), a`: an element passed ByRef beside the array to a procedure that erases or ReDims it. */
function checkElementPass(
	span: { start: number; end: number },
	dynamic: ReadonlySet<string>,
	procedures: ReadonlyMap<string, ProcedureNode | null>,
	source: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const call = extractCall(source, span);
	const callee = call && !call.qualifier ? procedures.get(call.name.toLowerCase()) : undefined;
	if (!call || !callee || call.slots.some(isNamedSlot)) {
		return;
	}
	const slots = call.slots.map((slot) => slot.filter((tok) => tok.kind !== 'comment'));
	slots.forEach((slot, k) => {
		const param = callee.params[k];
		const lower = tokenName(slot[0])?.toLowerCase();
		if (!param || !param.isArray || slot.length !== 1 || !lower || !dynamic.has(lower)) {
			return;
		}
		// Another argument is an element of the same array, passed ByRef.
		const element = slots.findIndex((other, j) => j !== k && tokenName(other[0])?.toLowerCase() === lower && other[1]?.rawText === '('
			&& other[other.length - 1]?.rawText === ')' && callee.params[j] !== undefined && !callee.params[j].byVal && !callee.params[j].isArray);
		if (element < 0) {
			return;
		}
		const name = param.name.toLowerCase();
		const unlocks = everyTime(source, callee.body, activity).some(({ toks }) => {
			const head = tokenText(toks[0]);
			const target = head === 'erase' ? toks[1] : head === 'redim' ? toks[tokenText(toks[1]) === 'preserve' ? 2 : 1] : undefined;
			return tokenName(target)?.toLowerCase() === name;
		});
		if (unlocks) {
			const at = slots[element][0];
			const offset = call.slotSpans?.[element]?.start;
			push(
				'arrayTemporarilyLocked',
				`'${call.name}' takes '${slots[element].map((tok) => tok.rawText).join('')}' ByRef, which locks '${slot[0].rawText}', and then resizes or erases that array through '${param.name}'. This will raise Run-time error '10': This array is fixed or temporarily locked.`,
				offset !== undefined ? { start: offset, end: offset + (slots[element][slots[element].length - 1].end - at.start) } : call.nameSpan,
			);
		}
	});
}
