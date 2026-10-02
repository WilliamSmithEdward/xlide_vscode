// Rule: a class that says `Implements IFoo` must supply every member of IFoo
// as `IFoo_Member`, with the interface's parameter list (issue #125).
// Measured in Excel 16.0 (build 20326, 2026-09-25):
//
//  - implements-member-missing: `Implements IFoo` in a class with no
//    `IFoo_Name` for IFoo's `Name` -> "Object module needs to implement
//    'Name' for interface 'IFoo'".
//  - implements-member-signature: `Private Function IFoo_Name(ByVal extra As
//    Long) As String` against `Public Function Name() As String` ->
//    "Procedure declaration does not match description of event or procedure
//    having the same name".
//
// The interface is read from the project index, so only a class module of
// this project is judged; an interface from a type library is not.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { ModuleNode } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { isProcedureKind, procedureParamsFromSymbol, type VbaProjectClassMember, type VbaProjectClassMembers, type VbaSymbol } from '../../symbols/symbolModel';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import { isObjectModuleKind, type PushFn } from '../analysisContext';
import { isKnownScalarType, normalizeType } from '../typeInference';
import {
	activeModuleMembers,
	absoluteSpan,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

export function checkImplementsMembers(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	moduleKind: ModuleSymbolKind,
	projectClassMembers: readonly VbaProjectClassMembers[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!isObjectModuleKind(moduleKind) || !projectClassMembers || projectClassMembers.length === 0) {
		return;
	}
	const procedures = new Map<string, VbaSymbol[]>();
	for (const symbol of symbols.root.children ?? []) {
		if (isProcedureKind(symbol.kind)) {
			const lower = symbol.name.toLowerCase();
			procedures.set(lower, [...(procedures.get(lower) ?? []), symbol]);
		}
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Statement') {
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, member.span);
		if (tokenText(toks[0]) !== 'implements') {
			continue;
		}
		// `Implements Lib.IFoo` names the interface last.
		const nameToken = [...toks].reverse().find((tok) => tokenName(tok) !== undefined);
		const interfaceName = nameToken ? tokenName(nameToken)! : undefined;
		if (!interfaceName || !nameToken) {
			continue;
		}
		const contract = projectClassMembers.find(
			(type) => type.kind === 'class' && type.exhaustive === true && type.name.toLowerCase() === interfaceName.toLowerCase(),
		);
		if (!contract) {
			continue;
		}
		for (const required of contract.members) {
			// A Friend member is no part of the interface (issue #291, measured).
			if (required.kind === 'event' || required.visibility === 'Friend') {
				continue;
			}
			const implementations = procedures.get(`${contract.name}_${required.name}`.toLowerCase()) ?? [];
			const variable = required.kind === 'property' && !required.procedureParams && implementations.length > 0
				? variableImplementationProblem(required, implementations)
				: undefined;
			if (variable) {
				push(variable.missing ? 'implementsMemberMissing' : 'implementsMemberSignature', variable.missing
					? `Object module needs to implement '${required.name}' for interface '${contract.name}': ${variable.message}.`
					: `'${variable.at.name}' does not match '${contract.name}.${required.name}': ${variable.message}. The procedure declaration must match the interface member it implements.`,
				variable.missing ? absoluteSpan(member.span, nameToken) : variable.at.nameSpan);
				continue;
			}
			if (implementations.length === 0) {
				push(
					'implementsMemberMissing',
					`Object module needs to implement '${required.name}' for interface '${contract.name}': add ${expectedProcedureLabel(contract.name, required)}.`,
					absoluteSpan(member.span, nameToken),
				);
				continue;
			}
			// A readable and writable property (a Public variable of the
			// interface, or a Get with a Let or Set) needs both accessors; Excel
			// refuses the project with the Get alone (issue #144, measured).
			if (required.kind === 'property' && required.writable && required.returns) {
				const hasGet = implementations.some((impl) => impl.kind === 'propertyGet');
				const hasSetter = implementations.some((impl) => impl.kind === 'propertyLet' || impl.kind === 'propertySet');
				if (!hasGet || !hasSetter) {
					push(
						'implementsMemberMissing',
						`Object module needs to implement '${required.name}' for interface '${contract.name}': add a Property ${hasGet ? 'Let or Set' : 'Get'} '${contract.name}_${required.name}' beside the Property ${hasGet ? 'Get' : 'Let'}.`,
						absoluteSpan(member.span, nameToken),
					);
				}
			}
			for (const implementation of implementations) {
				const problem = signatureMismatch(required, implementation) ?? passingMismatch(required, implementation);
				if (problem) {
					push(
						'implementsMemberSignature',
						`'${implementation.name}' does not match '${contract.name}.${required.name}': ${problem}. The procedure declaration must match the interface member it implements.`,
						implementation.nameSpan,
					);
				}
			}
		}
	}
}

/**
 * What else the VBE matches between an interface procedure and its
 * implementation (issue #291, measured in Excel 16.0): a Function is not
 * implemented by a Sub, and each parameter keeps its passing, ByVal or
 * ByRef (a plain one is ByRef), its Optional, and its default.
 */
function passingMismatch(required: VbaProjectClassMember, implementation: VbaSymbol): string | undefined {
	const declared = required.procedureParams;
	if (!declared) {
		return undefined;
	}
	if (declared.function && implementation.kind === 'sub') {
		return 'the interface member is a Function, and a Sub returns nothing';
	}
	const kind = implementation.kind as keyof typeof declared;
	const expected = declared[kind];
	if (!expected) {
		return undefined;
	}
	const actual = procedureParamsFromSymbol(implementation, { includePassing: true });
	for (let i = 0; i < Math.min(expected.length, actual.length); i++) {
		const want = expected[i];
		const got = actual[i];
		if (want.paramArray || got.paramArray) {
			continue;
		}
		if (Boolean(want.byVal) !== Boolean(got.byVal)) {
			return `parameter ${i + 1} is ${got.byVal ? 'ByVal' : 'ByRef'} here and ${want.byVal ? 'ByVal' : 'ByRef'} on the interface`;
		}
		if (want.optional !== got.optional) {
			return `parameter ${i + 1} is ${got.optional ? 'Optional' : 'required'} here and ${want.optional ? 'Optional' : 'required'} on the interface`;
		}
		if (want.optional && (want.defaultRaw ?? '').trim().toLowerCase() !== (got.defaultRaw ?? '').trim().toLowerCase()) {
			return `parameter ${i + 1} defaults to ${got.defaultRaw?.trim() || 'nothing'} here and ${want.defaultRaw?.trim() || 'nothing'} on the interface`;
		}
	}
	return undefined;
}

/**
 * A Public variable of the interface, implemented by Property procedures
 * (issue #291, measured in Excel 16.0). A value type needs a Get and a Let
 * whose value is ByVal and of the variable's type; an object type a Get
 * and a Set whose value is ByVal; a Variant a Get, a Let and a Set, neither
 * value ByVal.
 */
function variableImplementationProblem(required: VbaProjectClassMember, implementations: readonly VbaSymbol[]): { missing: boolean; message: string; at: VbaSymbol } | undefined {
	const type = normalizeType(required.writeType ?? required.returns) ?? 'variant';
	const object = type === 'object' || (type !== 'variant' && !isKnownScalarType(type));
	const byKind = (kind: string): VbaSymbol | undefined => implementations.find((impl) => impl.kind === kind);
	const get = byKind('propertyGet');
	const letter = byKind('propertyLet');
	const setter = byKind('propertySet');
	const name = `${implementations[0].name}`;
	const needs = (what: string): { missing: boolean; message: string; at: VbaSymbol } => ({ missing: true, message: `add a Property ${what} '${name}'`, at: implementations[0] });
	if (!get) {
		return needs('Get');
	}
	if (type === 'variant' ? !letter || !setter : object ? !setter : !letter) {
		return needs(type === 'variant' ? (letter ? 'Set' : 'Let') : object ? 'Set' : 'Let');
	}
	const valueOf = (procedure: VbaSymbol | undefined) => {
		const params = procedure ? procedureParamsFromSymbol(procedure, { includePassing: true }) : [];
		return params[params.length - 1];
	};
	for (const procedure of [letter, setter]) {
		const value = valueOf(procedure);
		if (!procedure || !value) {
			continue;
		}
		const byVal = Boolean(value.byVal);
		if (type === 'variant' ? byVal : !byVal) {
			return { missing: false, message: `its value is ${byVal ? 'ByVal' : 'ByRef'}, and a Public ${type === 'variant' ? 'Variant' : capitalize(required.writeType ?? type)} of the interface takes it ${byVal ? 'ByRef' : 'ByVal'}`, at: procedure };
		}
		if (!object && type !== 'variant' && (normalizeType(value.type) ?? 'variant') !== type) {
			return { missing: false, message: `its value is ${value.type ?? 'Variant'}, and the interface's variable is ${required.writeType}`, at: procedure };
		}
	}
	if (!object && type !== 'variant' && (normalizeType(get.asType) ?? 'variant') !== type) {
		return { missing: false, message: `it returns ${get.asType ?? 'Variant'}, and the interface's variable is ${required.writeType}`, at: get };
	}
	return undefined;
}

function capitalize(type: string): string {
	return type.charAt(0).toUpperCase() + type.slice(1);
}

function expectedProcedureLabel(interfaceName: string, member: VbaProjectClassMember): string {
	const name = `${interfaceName}_${member.name}`;
	if (member.kind === 'property') {
		return `a Property ${member.writable && !member.returns ? 'Let' : 'Get'} '${name}'`;
	}
	return `a ${member.returns ? 'Function' : 'Sub or Function'} '${name}'`;
}

interface ParsedParam {
	type: string;
	isArray: boolean;
}

/** The parameter list and return type an interface member's signature label states. */
function parseSignature(signature: string | undefined): { params: ParsedParam[]; returns: string } | undefined {
	if (!signature) {
		return undefined;
	}
	const open = signature.indexOf('(');
	if (open < 0) {
		return undefined;
	}
	// Parentheses and commas inside a string default are text: `Optional sep
	// As String = ", "` is one parameter and `= ")"` does not end the list
	// (issue #144).
	let depth = 0;
	let close = -1;
	let inString = false;
	const parts: string[] = [];
	let partStart = open + 1;
	for (let i = open; i < signature.length; i++) {
		const ch = signature[i];
		if (ch === '"') {
			inString = !inString;
			continue;
		}
		if (inString) {
			continue;
		}
		if (ch === '(') {
			depth++;
		} else if (ch === ')') {
			depth--;
			if (depth === 0) {
				close = i;
				break;
			}
		} else if (ch === ',' && depth === 1) {
			parts.push(signature.slice(partStart, i));
			partStart = i + 1;
		}
	}
	if (close < 0) {
		return undefined;
	}
	parts.push(signature.slice(partStart, close));
	const inner = signature.slice(open + 1, close).trim();
	const params = inner.length === 0 ? [] : parts.map((part) => {
		// Project signature labels wrap optional parameters in brackets. Those delimiters
		// are presentation, not part of the type (otherwise Date becomes "Date]").
		const label = part.trim();
		const parameter = label.startsWith('[') && label.endsWith(']') ? label.slice(1, -1) : label;
		const text = parameter.replace(/^(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)+/i, '').replace(/\s*=.*$/, '');
		const asMatch = /\sAs\s+(.+)$/i.exec(text);
		return {
			type: normalizeType(asMatch ? asMatch[1].trim() : undefined) ?? 'variant',
			isArray: /\(\s*\)/.test(text.replace(/\sAs\s.*$/i, '')),
		};
	});
	const returnsMatch = /\)\s*As\s+(.+)$/i.exec(signature.slice(close));
	return { params, returns: normalizeType(returnsMatch ? returnsMatch[1].trim() : undefined) ?? 'variant' };
}

function signatureMismatch(required: VbaProjectClassMember, implementation: VbaSymbol): string | undefined {
	const expected = parseSignature(required.signature);
	if (!expected) {
		return undefined;
	}
	const params = (implementation.children ?? []).filter((child) => child.kind === 'parameter');
	// A setter's last parameter is the value the interface property holds.
	const compared = implementation.kind === 'propertyLet' || implementation.kind === 'propertySet' ? params.slice(0, -1) : params;
	if (compared.length !== expected.params.length) {
		return `the interface member takes ${expected.params.length} parameter${expected.params.length === 1 ? '' : 's'}, this procedure ${compared.length}`;
	}
	for (let i = 0; i < compared.length; i++) {
		const actualType = normalizeType(compared[i].asType) ?? 'variant';
		if (actualType !== expected.params[i].type || Boolean(compared[i].isArray) !== expected.params[i].isArray) {
			return `parameter ${i + 1} is ${compared[i].asType ?? 'Variant'} here and ${expected.params[i].type} on the interface`;
		}
	}
	if (implementation.kind === 'function' || implementation.kind === 'propertyGet') {
		const actualReturn = normalizeType(implementation.asType) ?? 'variant';
		if (required.returns !== undefined && actualReturn !== expected.returns) {
			return `it returns ${implementation.asType ?? 'Variant'} where the interface member returns ${required.returns}`;
		}
	}
	return undefined;
}
