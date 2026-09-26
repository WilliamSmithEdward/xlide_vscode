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
import { isProcedureKind, type VbaProjectClassMember, type VbaProjectClassMembers, type VbaSymbol } from '../../symbols/symbolModel';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import { isObjectModuleKind, type PushFn } from '../analysisContext';
import { normalizeType } from '../typeInference';
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
			if (required.kind === 'event') {
				continue;
			}
			const implementations = procedures.get(`${contract.name}_${required.name}`.toLowerCase()) ?? [];
			if (implementations.length === 0) {
				push(
					'implementsMemberMissing',
					`Object module needs to implement '${required.name}' for interface '${contract.name}': add ${expectedProcedureLabel(contract.name, required)}.`,
					absoluteSpan(member.span, nameToken),
				);
				continue;
			}
			for (const implementation of implementations) {
				const problem = signatureMismatch(required, implementation);
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
	let depth = 0;
	let close = -1;
	for (let i = open; i < signature.length; i++) {
		if (signature[i] === '(') {
			depth++;
		} else if (signature[i] === ')') {
			depth--;
			if (depth === 0) {
				close = i;
				break;
			}
		}
	}
	if (close < 0) {
		return undefined;
	}
	const inner = signature.slice(open + 1, close).trim();
	const params = inner.length === 0 ? [] : inner.split(',').map((part) => {
		const text = part.trim().replace(/^(Optional|ByVal|ByRef|ParamArray)\s+/gi, '').replace(/\s*=.*$/, '');
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
