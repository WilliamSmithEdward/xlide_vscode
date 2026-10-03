// Rule family: module-level declarations that refer ahead or around a cycle
// (issue #211). Measured in 64-bit Excel 16.0 (build 20326, 2026-09-30):
//
//  - declaration-forward-reference: a Const, an Enum member, an array bound
//    or a `String * N` length that uses a Const or Enum member declared
//    further down its module, or itself: "Constant expression required". A
//    Type member of a Type declared further down: "Forward reference to
//    user-defined type". With the order swapped each compiles. A variable
//    of a later Type is fine, and so is order across modules.
//  - circular-declaration-dependency: a Type with a member of its own type,
//    and a cycle of Consts or Types through other modules: "Circular
//    dependencies between modules".

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { enumMemberRawExpression } from '../../constants/integerConstantExpression';
import { tokenize } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import type { VbaProjectClassMembers } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import {
	absoluteSpan,
	activeModuleMembers,
	isInactiveNode,
	matchParenFrom,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

/** Where each module-level constant and Type is declared; undefined for a name declared twice. */
interface Declarations {
	constants: Map<string, number | undefined>;
	/** Enum name -> member -> where the member is declared. */
	enums: Map<string, Map<string, number>>;
	types: Map<string, number | undefined>;
	/** The value expression of each constant and Enum member, for the cycle walk. */
	expressions: Map<string, string>;
}

export function checkDeclarationOrder(
	source: string,
	mod: ModuleNode,
	moduleName: string | undefined,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectClassMembers: readonly VbaProjectClassMembers[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const declared = moduleDeclarations(mod, activity);
	if (declared.constants.size === 0 && declared.types.size === 0) {
		return;
	}
	const check = (tokens: readonly VbaToken[], base: Span, siteStart: number, ownName: string | undefined): void => {
		for (const ref of constantReferences(tokens, declared)) {
			if (ref.start < siteStart) {
				continue;
			}
			const self = ref.start === siteStart && ref.name.toLowerCase() === ownName?.toLowerCase();
			push(
				'declarationForwardReference',
				self
					? `'${ref.name}' is defined in terms of itself. This is a VBE compile error: Constant expression required.`
					: `'${ref.name}' is declared further down the module, and a constant expression can only use what is declared above it. This is a VBE compile error: Constant expression required.`,
				absoluteSpan(base, ref.token),
			);
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			for (const decl of member.declarations) {
				if (isInactiveNode(activity, decl)) {
					continue;
				}
				const tokens = statementTokens(source, decl.span);
				check(
					member.isConst ? tokensAfterEquals(tokens) : declarationShapeTokens(tokens),
					decl.span,
					decl.span.start,
					member.isConst ? decl.name : undefined,
				);
			}
		} else if (member.kind === 'Enum') {
			for (const enumMember of member.members) {
				if (enumMember.valueRaw !== undefined && !isInactiveNode(activity, enumMember)) {
					check(tokensAfterEquals(statementTokens(source, enumMember.span)), enumMember.span, enumMember.span.start, enumMember.name);
				}
			}
		} else if (member.kind === 'Type') {
			const typeStart = declared.types.get(member.name.toLowerCase());
			for (const field of member.fields) {
				if (isInactiveNode(activity, field)) {
					continue;
				}
				const tokens = statementTokens(source, field.span);
				check(declarationShapeTokens(tokens), field.span, field.span.start, undefined);
				const asType = asTypeToken(tokens);
				const target = asType ? declared.types.get(tokenText(asType)) : undefined;
				if (!asType || target === undefined || typeStart === undefined || target < typeStart) {
					continue;
				}
				if (target === typeStart) {
					push(
						'circularDeclarationDependency',
						`Type '${member.name}' has a member of its own type. This is a VBE compile error: Circular dependencies between modules.`,
						absoluteSpan(field.span, asType),
					);
				} else {
					push(
						'declarationForwardReference',
						`Type '${asType.rawText}' is declared further down the module. This is a VBE compile error: Forward reference to user-defined type.`,
						absoluteSpan(field.span, asType),
					);
				}
			}
		}
	}

	if (moduleName !== undefined && projectIntegerConstants && projectIntegerConstants.size > 0) {
		checkConstantCyclesAcrossModules(mod, moduleName, declared, projectIntegerConstants, activity, push);
	}
	if (projectClassMembers?.some((surface) => surface.kind === 'userType' && surface.moduleName.toLowerCase() !== moduleName?.toLowerCase())) {
		checkTypeCyclesAcrossModules(source, mod, moduleName, projectClassMembers, activity, push);
	}
}

function moduleDeclarations(mod: ModuleNode, activity: ConditionalActivityTracker | undefined): Declarations {
	const constants = new Map<string, number | undefined>();
	const enums = new Map<string, Map<string, number>>();
	const types = new Map<string, number | undefined>();
	const expressions = new Map<string, string>();
	const place = (map: Map<string, number | undefined>, name: string, start: number): void => {
		const key = name.toLowerCase();
		map.set(key, map.has(key) ? undefined : start);
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && member.isConst) {
			for (const decl of member.declarations) {
				if (!isInactiveNode(activity, decl)) {
					place(constants, decl.name, decl.span.start);
					if (decl.defaultRaw !== undefined) {
						expressions.set(decl.name.toLowerCase(), decl.defaultRaw);
					}
				}
			}
		} else if (member.kind === 'Enum') {
			const members = new Map<string, number>();
			let previous: string | undefined;
			for (const enumMember of member.members) {
				if (isInactiveNode(activity, enumMember)) {
					continue;
				}
				place(constants, enumMember.name, enumMember.span.start);
				members.set(enumMember.name.toLowerCase(), enumMember.span.start);
				expressions.set(enumMember.name.toLowerCase(), enumMemberRawExpression(enumMember.valueRaw, previous));
				previous = enumMember.name;
			}
			enums.set(member.name.toLowerCase(), members);
		} else if (member.kind === 'Type') {
			place(types, member.name, member.span.start);
		}
	}
	return { constants, enums, types, expressions };
}

interface ConstantReference {
	name: string;
	token: VbaToken;
	start: number;
}

/**
 * The module's constants an expression names: a bare name, or an Enum member
 * qualified by its Enum. A name after any other `.` is some other object's
 * member and names nothing here.
 */
function constantReferences(tokens: readonly VbaToken[], declared: Declarations): ConstantReference[] {
	const out: ConstantReference[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const tok = tokens[i];
		const name = tokenName(tok);
		if (!name || tok.kind === 'bracketedIdentifier' || tokens[i - 1]?.rawText === '.') {
			continue;
		}
		if (tokens[i + 1]?.rawText === '.') {
			const member = tokens[i + 2];
			const start = member ? declared.enums.get(name.toLowerCase())?.get(tokenText(member)) : undefined;
			if (start !== undefined) {
				out.push({ name: `${name}.${member!.rawText}`, token: member!, start });
			}
			i += 2;
			continue;
		}
		const start = declared.constants.get(name.toLowerCase());
		if (start !== undefined) {
			out.push({ name, token: tok, start });
		}
	}
	return out;
}

/** The tokens after a declaration's top-level `=`: a Const's or Enum member's value. */
function tokensAfterEquals(tokens: readonly VbaToken[]): VbaToken[] {
	let depth = 0;
	for (let i = 0; i < tokens.length; i++) {
		const raw = tokens[i].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (depth === 0 && raw === '=' && tokens[i].kind === 'operator') {
			return tokens.slice(i + 1).filter((tok) => tok.kind !== 'comment');
		}
	}
	return [];
}

/**
 * The tokens a declaration's shape reads constants from: its array bounds and
 * the length of a `String * N`. The type after `As` is not among them.
 */
function declarationShapeTokens(tokens: readonly VbaToken[]): VbaToken[] {
	const out: VbaToken[] = [];
	const open = tokens.findIndex((tok) => tok.rawText === '(');
	const asAt = tokens.findIndex((tok) => tokenText(tok) === 'as');
	if (open >= 0 && (asAt < 0 || open < asAt)) {
		const close = matchParenFrom(tokens, open);
		out.push(...tokens.slice(open + 1, close < 0 ? tokens.length : close));
	}
	const star = tokens.findIndex((tok, i) => i > asAt && asAt >= 0 && tok.rawText === '*');
	if (star >= 0) {
		out.push(...tokens.slice(star + 1).filter((tok) => tok.kind !== 'comment'));
	}
	return out;
}

/** The type name after `As`, unless it is qualified (another module's) or a New. */
function asTypeToken(tokens: readonly VbaToken[]): VbaToken | undefined {
	const asAt = tokens.findIndex((tok) => tokenText(tok) === 'as');
	const name = asAt < 0 ? undefined : tokens[asAt + 1];
	if (!name || !tokenName(name) || tokens[asAt + 2]?.rawText === '.') {
		return undefined;
	}
	return name;
}

/**
 * Consts and Enum members that reach themselves through another module:
 * `Public Const A1 = B1 + 1` here and `Public Const B1 = A1 + 1` in Module2.
 * The other modules' values come from the project's exported constants,
 * keyed by bare and by qualified name; a value another module could fold is
 * already a number there and ends the walk.
 */
function checkConstantCyclesAcrossModules(
	mod: ModuleNode,
	moduleName: string,
	declared: Declarations,
	external: ReadonlyMap<string, string | undefined>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const own = moduleName.toLowerCase();
	// The module that exports each bare name, where exactly one does.
	const owners = new Map<string, string | undefined>();
	for (const key of external.keys()) {
		const dot = key.indexOf('.');
		if (dot > 0) {
			const bare = key.slice(dot + 1);
			const module = key.slice(0, dot);
			owners.set(bare, owners.has(bare) && owners.get(bare) !== module ? undefined : module);
		}
	}
	const modules = new Set([...owners.values()].filter((value): value is string => value !== undefined));

	interface Dependency { module: string; key: string; written: string }
	const prepared = new Map<string, Dependency[]>();
	const dependencies = (raw: string, scope: string): Dependency[] => {
		const cacheKey = scope + '\0' + raw;
		const cached = prepared.get(cacheKey);
		if (cached) { return cached; }
		const out: Dependency[] = [];
		const tokens = tokenize(raw).filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
		for (let i = 0; i < tokens.length; i++) {
			let name = tokenName(tokens[i]);
			if (!name || tokens[i - 1]?.rawText === '.') {
				continue;
			}
			let written = tokens[i].rawText;
			let qualifier: string | undefined;
			if (tokens[i + 1]?.rawText === '.' && modules.has(name.toLowerCase()) && tokenName(tokens[i + 2])) {
				qualifier = name.toLowerCase();
				name = tokenName(tokens[i + 2])!;
				written = `${tokens[i].rawText}.${tokens[i + 2].rawText}`;
				i += 2;
			}
			const lower = name.toLowerCase();
			let target: [string, string] | undefined;
			if (qualifier !== undefined) {
				target = qualifier === own ? [own, lower] : external.has(`${qualifier}.${lower}`) ? [qualifier, lower] : undefined;
			} else if (scope === own) {
				target = declared.expressions.has(lower) ? [own, lower] : owners.get(lower) ? [owners.get(lower)!, lower] : undefined;
			} else if (external.has(`${scope}.${lower}`)) {
				target = [scope, lower];
			} else if (declared.expressions.has(lower)) {
				target = [own, lower];
			} else if (owners.get(lower)) {
				target = [owners.get(lower)!, lower];
			}
			if (!target) {
				continue;
			}
			out.push({ module: target[0], key: target[1], written });
		}
		prepared.set(cacheKey, out);
		return out;
	};

	const reachesItself = (start: string): string | undefined => {
		const raw = declared.expressions.get(start);
		if (raw === undefined) { return undefined; }
		const seen = new Set<string>();
		// Explicit DFS frames preserve reference order and the first foreign name
		// without consuming the JavaScript call stack for long dependency chains.
		const stack: Array<{ edges: Dependency[]; index: number; via: string | undefined }> = [
			{ edges: dependencies(raw, own), index: 0, via: undefined },
		];
		while (stack.length > 0) {
			const frame = stack[stack.length - 1];
			if (frame.index === frame.edges.length) { stack.pop(); continue; }
			const { module, key, written } = frame.edges[frame.index++];
			if (module === own && key === start && frame.via !== undefined) { return frame.via; }
			const id = module + '.' + key;
			if (seen.has(id)) { continue; }
			seen.add(id);
			const next = module === own ? declared.expressions.get(key) : external.get(id);
			if (next === undefined || /^[-+]?\d+(\.\d+)?$/.test(next.trim())) { continue; }
			stack.push({
				edges: dependencies(next, module), index: 0,
				via: frame.via ?? (module === own ? undefined : written),
			});
		}
		return undefined;
	};

	for (const member of activeModuleMembers(mod, activity)) {
		const names: { name: string; span: Span }[] = [];
		if (member.kind === 'VariableGroup' && member.isConst) {
			for (const decl of member.declarations) {
				if (!isInactiveNode(activity, decl)) {
					names.push({ name: decl.name, span: decl.nameSpan ?? decl.span });
				}
			}
		} else if (member.kind === 'Enum') {
			for (const enumMember of member.members) {
				if (!isInactiveNode(activity, enumMember)) {
					names.push({ name: enumMember.name, span: enumMember.nameSpan ?? enumMember.span });
				}
			}
		}
		for (const { name, span } of names) {
			const via = reachesItself(name.toLowerCase());
			if (via) {
				push(
					'circularDeclarationDependency',
					`'${name}' depends on itself through '${via}' in another module. This is a VBE compile error: Circular dependencies between modules.`,
					span,
				);
			}
		}
	}
}

/**
 * Types that contain themselves through another module's Type:
 * `Public Type T1: a As T2` here and `Public Type T2: b As T1` in Module2.
 */
function checkTypeCyclesAcrossModules(
	source: string,
	mod: ModuleNode,
	moduleName: string | undefined,
	projectClassMembers: readonly VbaProjectClassMembers[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const own = moduleName?.toLowerCase();
	const byName = new Map<string, VbaProjectClassMembers | undefined>();
	for (const surface of projectClassMembers) {
		if (surface.kind !== 'userType') {
			continue;
		}
		const key = surface.name.toLowerCase();
		byName.set(key, byName.has(key) ? undefined : surface);
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Type') {
			continue;
		}
		const start = member.name.toLowerCase();
		for (const field of member.fields) {
			if (isInactiveNode(activity, field)) {
				continue;
			}
			const asType = asTypeToken(statementTokens(source, field.span));
			const first = asType ? byName.get(tokenText(asType)) : undefined;
			if (!asType || !first || first.moduleName.toLowerCase() === own) {
				continue;
			}
			// byName resolves each unambiguous name to one canonical surface.
			// Mark it when queued so shared descendants enter only once.
			const seen = new Set<VbaProjectClassMembers>([first]);
			const queue: VbaProjectClassMembers[] = [first];
			let cycle = false;
			for (let head = 0; head < queue.length && !cycle; head++) {
				const surface = queue[head];
				for (const fieldMember of surface.members) {
					const typeName = fieldMember.returns?.toLowerCase();
					if (!typeName) {
						continue;
					}
					const next = byName.get(typeName);
					if (typeName === start && (next === undefined || next.moduleName.toLowerCase() === own)) {
						cycle = true;
						break;
					}
					if (next && next.moduleName.toLowerCase() !== own && !seen.has(next)) {
						seen.add(next);
						queue.push(next);
					}
				}
			}
			if (cycle) {
				push(
					'circularDeclarationDependency',
					`Type '${member.name}' contains itself through ${first.moduleName}.${first.name}. This is a VBE compile error: Circular dependencies between modules.`,
					absoluteSpan(field.span, asType),
				);
			}
		}
	}
}
