// Value completion for an argument whose parameter declares an enumeration.
//
// `ThisWorkbook.BreakLink Name:="test", Type:=` has exactly two legal values,
// and the model knows both: the signature says `Type As XlLinkType` and the
// enumeration carries xlLinkTypeExcelLinks and xlLinkTypeOLELinks. Without this
// the caret sat at `:=` and got the general identifier list - every global,
// every constant in the library, in no useful order.
//
// The active parameter comes from the same resolver the call tip uses, so the
// list and the tip can never disagree about which parameter is being typed.
// A NAMED argument is matched by its name rather than by position, because
// `Type:=` says which parameter it is no matter where it sits.
//
// Pure analyzer code: no `vscode` dependency.

import { tokenizeCached } from '../lexer/tokenize';
import { isIdentLike } from '../lexer/tokenHelpers';
import { getHostEnumMembers, resolveHostEnum } from '../host/hostModel';
import type { HostConstant } from '../host/excelObjectModel';
import { resolveSignatureHelp, type SignatureHelpContext } from '../signature/signatureHelp';

/** The enumeration a caret's argument slot accepts, and its members. */
export interface ArgumentValueCompletion {
	/** Enumeration name as declared, e.g. "XlLinkType". */
	enumName: string;
	/** The enumeration's constants, in declaration order. */
	constants: readonly HostConstant[];
	/** Parameter this came from, as written in the signature. */
	parameter: string;
}

/** Type name from a parameter label: "Type As XlLinkType" -> "XlLinkType". */
function parameterTypeName(label: string): string | undefined {
	return label.replace(/[[\]]/g, '').trim().match(/\bAs\s+([A-Za-z_][A-Za-z0-9_.]*)$/)?.[1];
}

/**
 * The parameter name a named argument at `offset` addresses, when the caret sits
 * after `Name:=`. Undefined for a positional slot, which falls back to the call
 * tip's active parameter.
 */
function namedArgumentAt(source: string, offset: number): string | undefined {
	const tokens = tokenizeCached(source);
	// Last two significant tokens ending at or before the caret, skipping the
	// partial value already typed.
	let i = tokens.length - 1;
	while (i >= 0 && tokens[i].start >= offset) {
		i -= 1;
	}
	if (i >= 0 && isIdentLike(tokens[i]) && tokens[i].end <= offset) {
		i -= 1;   // a value being typed: `Type:=xlLink`
	}
	if (i < 1 || tokens[i].rawText !== ':=') {
		return undefined;
	}
	const name = tokens[i - 1];
	return isIdentLike(name) || name.kind === 'bracketedIdentifier'
		? name.rawText.replace(/[[\]]/g, '')
		: undefined;
}

/**
 * The enumeration accepted at `offset`, or undefined when the caret is not in an
 * argument slot, the parameter declares no enumeration, or the host model does
 * not carry it.
 */
export function resolveArgumentValueCompletion(
	source: string,
	offset: number,
	ctx: SignatureHelpContext = {},
): ArgumentValueCompletion | undefined {
	const info = resolveSignatureHelp(source, offset, ctx);
	if (!info || info.parameters.length === 0) {
		return undefined;
	}
	const named = namedArgumentAt(source, offset);
	const label = named
		? info.parameters.find(
			(parameter) => parameter.label
				.replace(/[[\]]/g, '')
				.trim()
				.split(/\s+/)[0]
				?.toLowerCase() === named.toLowerCase(),
		)?.label
		: info.parameters[info.activeParameter]?.label;
	if (!label) {
		return undefined;
	}
	const typeName = parameterTypeName(label);
	if (!typeName) {
		return undefined;
	}
	const hostEnum = resolveHostEnum(typeName, ctx.model);
	if (!hostEnum) {
		return undefined;
	}
	const constants = getHostEnumMembers(hostEnum.displayName, ctx.model);
	return constants.length > 0
		? { enumName: hostEnum.displayName, constants, parameter: label }
		: undefined;
}
