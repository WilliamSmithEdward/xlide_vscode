// Rule family: late binding.
//
// A receiver whose static type is Variant or Object dispatches through
// IDispatch at runtime. Friend and Private members are not on a class's
// dispatch interface, so a late-bound call to one raises runtime error 438 -
// and the VBA compiler says nothing, because it cannot know the runtime type
// either. That combination makes the failure easy to ship: the code compiles,
// and the call only dies on the first execution that reaches it.

import type { VbaToken } from '../../lexer/tokenKinds';
import type { Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProjectClassMembers, VbaSymbol } from '../../symbols/symbolModel';
import { isHostMemberName } from '../../host/hostModel';
import { resolveRuntimeObject, resolveRuntimeObjectType } from '../../runtime/vbaRuntime';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import {
	declarationShapeEnvironmentFor,
	declaredValueTypeForSourceBinding,
	normalizeType,
} from '../typeInference';
import {
	statementTokens,
	tokenName,
	type ProcedureStatementVisitor,
} from '../walker';

/** Declared types whose members can only be reached through IDispatch. */
const LATE_BOUND_TYPES = new Set(['variant', 'object']);

/**
 * Per-statement rule: member access on a late-bound receiver where the member
 * name is Friend-only across the project's class modules.
 *
 * Deliberately narrow. Project class members that are `Private` never reach
 * `projectClassMembers` at all (the index drops them), so "unknown member" and
 * "Private member" are indistinguishable here - and the VBE oracle says an
 * unknown member on a Variant/Object receiver must stay silent, because the
 * runtime type is unknowable and the call may well be legal. Firing only on
 * names that resolve EXCLUSIVELY to Friend members keeps the rule on the one
 * case where there is provably no legal late-bound target anywhere in scope.
 */
export function checkLateBoundFriendMember(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	projectClassMembers: readonly VbaProjectClassMembers[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const friendOnly = friendOnlyMemberNames(projectClassMembers);
	if (friendOnly.size === 0) {
		return () => undefined;
	}
	return (member) => {
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		const declaredTypeOf = (name: string): string | undefined => {
			const shape = shapes.get(name.toLowerCase());
			if (shape) {
				// An entry with no `As` clause is an implicit Variant.
				return shape.asType ?? 'Variant';
			}
			const binding = declaredValueTypeForSourceBinding(
				symbols, procSym, projectVisibleSymbols, name,
			);
			// `resolved` with no type also covers ambiguous multi-definition
			// bindings, which are not proof of anything - stay quiet.
			return binding.resolved ? binding.asType : undefined;
		};
		return (stmt) => {
			const toks = statementTokens(source, stmt.span);
			for (const access of lateBoundMemberAccesses(toks, declaredTypeOf)) {
				const owners = friendOnly.get(access.member.toLowerCase());
				if (!owners) { continue; }
				push(
					'lateBoundFriendMember',
					`'${access.member}' is a Friend member of ${describeOwners(owners)}, so it is not on the `
					+ `dispatch interface. ${access.receiverDescription} is late bound (${access.receiverType}), `
					+ 'so this raises run-time error 438 at run time even though it compiles. '
					+ 'Assign to a variable of the class type first, then access the member through that.',
					{ start: stmt.span.start + access.token.start, end: stmt.span.start + access.token.end },
				);
			}
		};
	};
}

interface LateBoundAccess {
	member: string;
	token: VbaToken;
	receiverType: string;
	receiverDescription: string;
}

/**
 * Find `<late-bound receiver>.Member` in one statement's tokens.
 *
 * Two receiver shapes are recognized, both of which appear in real code:
 *   - a bare identifier declared As Variant/As Object/with no type; and
 *   - a `Collection` subscript or `.Item(...)` call, whose result is Variant.
 *     This is the shape that hides the bug in practice, because the collection
 *     itself is strongly typed and only its element type is lost.
 */
function lateBoundMemberAccesses(
	toks: readonly VbaToken[],
	declaredTypeOf: (name: string) => string | undefined,
): LateBoundAccess[] {
	const out: LateBoundAccess[] = [];
	for (let i = 1; i < toks.length - 1; i++) {
		if (toks[i].rawText !== '.') { continue; }
		const memberName = tokenName(toks[i + 1]);
		if (!memberName) { continue; }

		const receiver = receiverBefore(toks, i, declaredTypeOf);
		if (!receiver) { continue; }
		out.push({
			member: memberName,
			token: toks[i + 1],
			receiverType: receiver.type,
			receiverDescription: receiver.description,
		});
	}
	return out;
}

function receiverBefore(
	toks: readonly VbaToken[],
	dotIndex: number,
	declaredTypeOf: (name: string) => string | undefined,
): { type: string; description: string } | undefined {
	const prev = toks[dotIndex - 1];
	if (!prev) { return undefined; }

	if (prev.rawText === ')') {
		const open = matchingOpenParen(toks, dotIndex - 1);
		if (open === undefined) { return undefined; }
		return collectionElementReceiver(toks, open, declaredTypeOf);
	}

	// A bare identifier, and not itself the tail of a longer chain: a chain like
	// `a.b.c` would need b's return type, which is a different question.
	if (toks[dotIndex - 2]?.rawText === '.') { return undefined; }
	const name = tokenName(prev);
	if (!name) { return undefined; }
	const normalized = normalizeType(declaredTypeOf(name));
	if (!normalized || !LATE_BOUND_TYPES.has(normalized)) { return undefined; }
	// A name that is also a built-in runtime object (Err, Debug) is not a
	// late-bound local no matter what a same-named declaration says.
	if (resolveRuntimeObject(name)) { return undefined; }
	return {
		type: normalized === 'object' ? 'As Object' : 'As Variant',
		description: `'${name}'`,
	};
}

/**
 * Recognize `coll(i)` and `coll.Item(i)` where `coll` is declared As Collection.
 * `Collection.Item` returns Variant, so the element arrives late bound however
 * strongly typed the objects inside the collection actually are.
 */
function collectionElementReceiver(
	toks: readonly VbaToken[],
	openIndex: number,
	declaredTypeOf: (name: string) => string | undefined,
): { type: string; description: string } | undefined {
	let rootIndex = openIndex - 1;
	let viaItem = false;
	const beforeParen = tokenName(toks[rootIndex]);
	if (beforeParen && beforeParen.toLowerCase() === 'item' && toks[rootIndex - 1]?.rawText === '.') {
		viaItem = true;
		rootIndex -= 2;
	}
	if (toks[rootIndex - 1]?.rawText === '.') { return undefined; }
	const rootName = tokenName(toks[rootIndex]);
	if (!rootName) { return undefined; }
	if (normalizeType(declaredTypeOf(rootName)) !== 'collection') { return undefined; }
	return {
		type: 'Collection.Item returns Variant',
		description: viaItem ? `'${rootName}.Item(...)'` : `'${rootName}(...)'`,
	};
}

function matchingOpenParen(toks: readonly VbaToken[], closeIndex: number): number | undefined {
	let depth = 0;
	for (let i = closeIndex; i >= 0; i--) {
		const text = toks[i].rawText;
		if (text === ')') { depth++; }
		else if (text === '(') {
			depth--;
			if (depth === 0) { return i; }
		}
	}
	return undefined;
}

/**
 * Member names that appear ONLY as Friend members of project class modules.
 *
 * A name that is also Public anywhere, or that exists in the host object model
 * or the VBA runtime objects, could legally dispatch to that other target, so
 * it is excluded - the runtime type inside a Variant is unknowable and the rule
 * must not guess.
 */
function friendOnlyMemberNames(
	projectClassMembers: readonly VbaProjectClassMembers[] | undefined,
): Map<string, string[]> {
	const friendOwners = new Map<string, string[]>();
	const disqualified = new Set<string>();
	for (const owner of projectClassMembers ?? []) {
		// Only plain class modules are source-exhaustive. Document modules and
		// UserForms also expose host/designer members we cannot enumerate.
		const exhaustiveClass = owner.kind === 'class' && owner.exhaustive !== false;
		for (const member of owner.members) {
			const lower = member.name.toLowerCase();
			if (member.visibility === 'Friend' && exhaustiveClass) {
				const owners = friendOwners.get(lower);
				if (owners) { owners.push(owner.name); } else { friendOwners.set(lower, [owner.name]); }
			} else {
				disqualified.add(lower);
			}
		}
	}
	const out = new Map<string, string[]>();
	for (const [lower, owners] of friendOwners) {
		if (disqualified.has(lower)) { continue; }
		if (isHostMemberName(lower)) { continue; }
		if (isRuntimeObjectMemberName(lower)) { continue; }
		out.set(lower, owners);
	}
	return out;
}

const RUNTIME_OBJECT_MEMBER_NAMES = new Set<string>();
function isRuntimeObjectMemberName(lower: string): boolean {
	if (RUNTIME_OBJECT_MEMBER_NAMES.size === 0) {
		for (const name of ['Collection', 'Err', 'Debug', 'Dictionary']) {
			const object = resolveRuntimeObject(name) ?? resolveRuntimeObjectType(name);
			for (const member of object?.members ?? []) {
				RUNTIME_OBJECT_MEMBER_NAMES.add(member.name.toLowerCase());
			}
		}
	}
	return RUNTIME_OBJECT_MEMBER_NAMES.has(lower);
}

function describeOwners(owners: readonly string[]): string {
	const unique = [...new Set(owners)];
	if (unique.length === 1) { return `class '${unique[0]}'`; }
	return `classes ${unique.map((name) => `'${name}'`).join(', ')}`;
}
