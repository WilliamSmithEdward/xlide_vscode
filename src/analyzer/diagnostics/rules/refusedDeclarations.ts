// Rule family: declarations the VBE refuses by their form (issue #212).
// Measured in 64-bit Excel 16.0 (build 20326, 2026-09-30), by compiling the
// whole project:
//
//  - private-type-in-public-signature: a Private Enum of a standard module
//    in a Public (or unmarked) Sub, Function, Property, Declare or variable;
//    in a class module a Private Enum or Type in a Public procedure or Event,
//    and a Private Enum as a Public variable. A Private Type is fine in a
//    standard module's public signatures, and a Public Type's field may be of
//    a Private Enum.
//  - optional-property-value: a Property Let or Set whose value parameter is
//    Optional, "Syntax error". An Optional index before it compiles.
//  - event-parameter-form: an Event parameter that is Optional or a
//    ParamArray, "Syntax error".
//  - const-invalid-type: a Const declared As a type that is not one of VBA's
//    own: As Object is "Invalid data type for constant", and As Collection,
//    Range, a class, an Enum, a Type or Decimal is "Expected: type name".
//  - empty-enum: an Enum with no members, "Enum without members not allowed".
//  - type-member-without-type: a Type member with no As clause, a type
//    suffix included, "Statement invalid inside Type block".
//  - type-enum-name-conflict: a Type and an Enum of one name in one module,
//    "Ambiguous name detected". A Const may share either's name.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { ModuleNode, ParameterNode, Span, VariableDeclNode } from '../../parser/nodes';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import {
	activeModuleMembers,
	declaredNameSpan,
	forEachVariableGroup,
	isInactiveNode,
	matchParenFrom,
	statementTokens,
	tokenText,
} from '../walker';

const PRIVATE_TYPE_MESSAGE =
	'Private Enum and user defined types cannot be used as parameters or return types for public procedures, public data members, or fields of public user defined types.';

/**
 * The types a Const may be declared As, and Decimal, which is refused here as
 * well but is invalid-as-type-name's to report wherever it appears.
 */
const CONST_TYPES = new Set([
	'boolean', 'byte', 'integer', 'long', 'longlong', 'longptr', 'currency', 'single', 'double', 'date', 'string', 'variant',
	'decimal',
]);

export function checkRefusedDeclarations(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	checkPrivateTypesInPublicSignatures(source, mod, moduleKind, activity, push);
	const typeNames = new Map<string, Span>();
	const enumNames = new Map<string, Span>();
	// Two Enums or two Types of one name: Ambiguous name detected (issue
	// #639, measured in Excel 16.0).
	const repeatedName = (seen: Map<string, Span>, name: string, span: Span, kinds: string): void => {
		if (seen.has(name.toLowerCase())) {
			push('typeEnumNameConflict', `Two ${kinds} in this module are both named '${name}'. This is a VBE compile error: Ambiguous name detected.`, span);
		} else {
			seen.set(name.toLowerCase(), span);
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		switch (member.kind) {
			case 'Procedure':
				if ((member.procKind === 'PropertyLet' || member.procKind === 'PropertySet') && member.params.at(-1)?.optional) {
					const value = member.params.at(-1)!;
					push(
						'optionalPropertyValue',
						`The value parameter '${value.name}' of a Property ${member.procKind === 'PropertyLet' ? 'Let' : 'Set'} cannot be Optional. This is a VBE compile error: Syntax error.`,
						value.nameSpan ?? value.span,
					);
				}
				forEachVariableGroup(member.body, (group) => {
					if (group.isConst) {
						checkConstTypes(source, group.declarations, activity, push);
					}
				}, activity);
				break;
			case 'Event': {
				for (const param of member.params) {
					if (param.optional || param.paramArray) {
						push(
							'eventParameterForm',
							`An Event parameter cannot be ${param.paramArray ? 'a ParamArray' : 'Optional'}: '${param.name}' in '${member.name}'. This is a VBE compile error: Syntax error.`,
							param.nameSpan ?? param.span,
						);
					} else if (param.isArray && param.byVal) {
						// An array passes ByRef (issue #266, measured).
						push(
							'eventParameterForm',
							`An Event's array parameter must be ByRef: '${param.name}' in '${member.name}'. This is a VBE compile error: Array argument must be ByRef.`,
							param.nameSpan ?? param.span,
						);
					}
				}
				// An Event is Public: `Private Event` and `Friend Event` are
				// "Expected: Sub or Function or Property", and an As clause after
				// it is "Expected: end of statement" (issue #266, measured in a
				// class). In a standard module any Event is
				// event-declaration-module-kind's.
				if (moduleKind === 'standard') {
					break;
				}
				const toks = statementTokens(source, member.span);
				const head = tokenText(toks[0]);
				if (head === 'private' || head === 'friend') {
					push(
						'invalidProcedureHeader',
						`An Event cannot be ${toks[0].rawText}; an Event is always Public. This is a VBE compile error: Expected: Sub or Function or Property.`,
						{ start: member.span.start + toks[0].start, end: member.span.start + toks[0].end },
					);
				}
				const open = toks.findIndex((tok) => tok.rawText === '(');
				const close = open < 0 ? -1 : matchParenFrom(toks, open);
				const after = close < 0 ? undefined : toks[close + 1];
				if (after && tokenText(after) === 'as') {
					push(
						'invalidProcedureHeader',
						`An Event returns nothing, so it takes no As clause: '${member.name}'. This is a VBE compile error: Expected: end of statement.`,
						{ start: member.span.start + after.start, end: member.span.start + toks[toks.length - 1].end },
					);
				}
				break;
			}
			case 'VariableGroup':
				if (member.isConst) {
					checkConstTypes(source, member.declarations, activity, push);
				}
				// Static keeps a local's value between calls, and means nothing
				// outside a procedure (issue #216).
				if (member.modifier.toLowerCase() === 'static') {
					push(
						'staticOutsideProcedure',
						'Static declares a variable inside a procedure; at module level use Private or Dim. This is a VBE compile error: Invalid outside procedure.',
						{ start: member.span.start, end: member.span.start + 'Static'.length },
					);
				}
				break;
			case 'Enum':
				if (member.closed && !member.members.some((enumMember) => !isInactiveNode(activity, enumMember))) {
					push(
						'emptyEnum',
						`Enum '${member.name}' must declare at least one member. This is a VBE compile error: Enum without members not allowed.`,
						member.nameSpan ?? member.span,
					);
				}
				repeatedName(enumNames, member.name, member.nameSpan ?? member.span, 'Enums');
				break;
			case 'Type':
				for (const field of member.fields) {
					if (!field.hasAsClause && !isInactiveNode(activity, field)) {
						// `10  Id As Long` reads as a member named 10: a line
						// number, which a Type block does not take either.
						push(
							'typeMemberWithoutType',
							/^\d+$/.test(field.name)
								? `A line number cannot label a Type member. This is a VBE compile error: Statement invalid inside Type block.`
								: `Type member '${field.name}' needs an As clause. This is a VBE compile error: Statement invalid inside Type block.`,
							field.nameSpan ?? field.span,
						);
					}
				}
				repeatedName(typeNames, member.name, member.nameSpan ?? member.span, 'Types');
				break;
			default:
				break;
		}
	}
	for (const [name, span] of enumNames) {
		const typeSpan = typeNames.get(name);
		if (typeSpan) {
			const later = typeSpan.start > span.start ? typeSpan : span;
			push(
				'typeEnumNameConflict',
				`A Type and an Enum in this module are both named '${source.slice(later.start, later.end)}'. This is a VBE compile error: Ambiguous name detected.`,
				later,
			);
		}
	}
}

function checkConstTypes(
	source: string,
	declarations: readonly VariableDeclNode[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const decl of declarations) {
		const asType = decl.asType?.trim();
		if (!asType || isInactiveNode(activity, decl) || CONST_TYPES.has(asType.toLowerCase())) {
			continue;
		}
		const object = asType.toLowerCase() === 'object';
		push(
			'constInvalidType',
			object
				? `A Const cannot be declared As Object. This is a VBE compile error: Invalid data type for constant.`
				: `A Const can only be declared As one of VBA's own types, not As ${asType}. This is a VBE compile error: Expected: type name.`,
			declaredNameSpan(source, decl.span, decl.name),
		);
	}
}

/** Whether a procedure-like member is reachable from outside its module. */
function isPublicMember(modifiers: readonly string[]): boolean {
	return !modifiers.some((modifier) => /^(private|friend)$/i.test(modifier));
}

function checkPrivateTypesInPublicSignatures(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (moduleKind !== 'standard' && moduleKind !== 'class') {
		return;
	}
	const privateEnums = new Set<string>();
	const privateTypes = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Enum' && member.visibility?.toLowerCase() === 'private') {
			privateEnums.add(member.name.toLowerCase());
		} else if (member.kind === 'Type' && member.visibility?.toLowerCase() === 'private') {
			privateTypes.add(member.name.toLowerCase());
		}
	}
	// A standard module may use its Private Types publicly; a class may not.
	const refused = (typeName: string | undefined): boolean => {
		const key = typeName?.replace(/\(\s*\)$/, '').trim().toLowerCase();
		return key !== undefined && (privateEnums.has(key) || (moduleKind === 'class' && privateTypes.has(key)));
	};
	if (privateEnums.size === 0 && (moduleKind !== 'class' || privateTypes.size === 0)) {
		return;
	}
	const report = (span: Span): void => push('privateTypeInPublicSignature', `${PRIVATE_TYPE_MESSAGE} This is a VBE compile error.`, span);
	const checkParams = (params: readonly ParameterNode[]): void => {
		for (const param of params) {
			if (refused(param.asType)) {
				report(param.nameSpan ?? param.span);
			}
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && isPublicMember(member.modifiers)) {
			checkParams(member.params);
			if (refused(member.returnType)) {
				report(member.nameSpan ?? member.span);
			}
		} else if (member.kind === 'Declare' && isPublicMember([member.visibility ?? ''])) {
			checkParams(member.params);
			if (refused(member.returnType)) {
				report(member.nameSpan ?? member.span);
			}
		} else if (member.kind === 'Event' && isPublicMember([member.visibility ?? ''])) {
			checkParams(member.params);
		} else if (member.kind === 'VariableGroup' && !member.isConst && /^(public|global)$/i.test(member.modifier)) {
			for (const decl of member.declarations) {
				// A class's Public variable of its Private Type is the object-module
				// rule's; only the Enum is this message there.
				const key = decl.asType?.trim().toLowerCase();
				if (key !== undefined && privateEnums.has(key) && !isInactiveNode(activity, decl)) {
					report(declaredNameSpan(source, decl.span, decl.name));
				}
			}
		}
	}
}
