// Procedure names in strings (issue #217).
//
// A string can name a procedure that runs later: `Application.Run
// "Module1.DoIt"`, `Application.OnTime Now, "Tick"`, `shp.OnAction =
// "Sheet1.Clicked"`, and a framework's own wiring, ReDim's
// `ui.Button("run").OnClick "Demo.BuildReport"`, which hands the string to
// Application.Run when the button is clicked. Such a string is offered the
// project's procedures, and hover and go-to-definition follow it to the one
// it names.
//
// A string is read as a procedure name where the parameter it fills is
// named for one: Macro (Application.Run), Procedure (Application.OnTime),
// or any name ending in Proc, Procedure or Macro (ReDim's handlerProc,
// buildProc and checkProc). OnAction is a property, assigned with `=`.
//
// Pure analyzer code: no `vscode` dependency.

import { tokenizeCached } from '../lexer/tokenize';
import { firstTokenEndingAtOrAfter } from '../lexer/tokenHelpers';
import type { VbaToken } from '../lexer/tokenKinds';
import type { Span } from '../parser/nodes';
import type { VbaProcedureSignature } from '../symbols/symbolModel';
import { resolveSignatureHelp, type SignatureHelpContext } from '../signature/signatureHelp';
import { stringLiteralValue } from '../diagnostics/typeInference';

/** A string that names a procedure: its text, and the span of what is between its quotes. */
export interface MacroNameString {
	text: string;
	/** The characters between the quotes, from the opening quote's right to the closing one or the caret. */
	contentSpan: Span;
}

/** A procedure a macro-name string can name. */
export interface MacroNameCandidate {
	/** `Module.Proc`, as the string spells it. */
	name: string;
	procedure: VbaProcedureSignature;
}

const MACRO_PARAMETER = /^(macro|procedure|\w*(proc|procedure|macro))$/i;

/** The parameter name a signature label declares: `[ByVal handlerProc As String]` -> `handlerProc`. */
function parameterName(label: string): string | undefined {
	const words = label.replace(/[[\]]/g, ' ').trim().split(/\s+/);
	const name = words.find((word) => !/^(optional|byval|byref|paramarray)$/i.test(word));
	return name?.replace(/[^A-Za-z0-9_].*$/, '') || undefined;
}

/** The token before index `i` on the same logical line. */
function previous(tokens: readonly VbaToken[], i: number): VbaToken | undefined {
	const tok = tokens[i - 1];
	return tok && tok.kind !== 'newline' && tok.kind !== 'colon' ? tok : undefined;
}

/**
 * The procedure-name string at `offset`, the caret inside it or on it, or
 * undefined where the string fills no parameter named for a procedure.
 */
export function macroNameStringAt(source: string, offset: number, ctx: SignatureHelpContext = {}): MacroNameString | undefined {
	const tokens = tokenizeCached(source);
	const index = firstTokenEndingAtOrAfter(tokens, offset);
	const token = tokens[index];
	if (!token || token.kind !== 'stringLiteral' || !(offset > token.start && offset <= token.end)) {
		return undefined;
	}
	// Empty strings and strings ending in an escaped quote still have a closing delimiter.
	const closed = /^"(?:[^"]|"")*"$/.test(token.rawText);
	const contentSpan: Span = { start: token.start + 1, end: closed ? token.end - 1 : token.end };
	const text = closed ? stringLiteralValue(token.rawText) : token.rawText.slice(1).replace(/""/g, '"');
	const before = previous(tokens, index);
	const word = before?.rawText.toLowerCase();
	// `shp.OnAction = "Proc"` and `.OnAction = "Proc"`.
	if (word === '=') {
		const target = previous(tokens, index - 1);
		return target?.rawText.toLowerCase() === 'onaction' ? { text, contentSpan } : undefined;
	}
	// `handlerProc:="Proc"`, a named argument.
	if (word === ':=') {
		const named = previous(tokens, index - 1);
		return named && MACRO_PARAMETER.test(named.rawText) ? { text, contentSpan } : undefined;
	}
	const help = resolveSignatureHelp(source, token.start + 1, ctx);
	const label = help?.parameters[help.activeParameter]?.label;
	const name = label ? parameterName(label) : undefined;
	return name && MACRO_PARAMETER.test(name) ? { text, contentSpan } : undefined;
}

/** Reject definite non-macro string positions before building editor project facts.
 * Positional call arguments remain possible until their signature is known. */
export function macroNameStringMayResolveAt(source: string, offset: number): boolean {
	const tokens = tokenizeCached(source);
	const index = firstTokenEndingAtOrAfter(tokens, offset);
	const token = tokens[index];
	if (!token || token.kind !== 'stringLiteral' || offset <= token.start || offset > token.end) {
		return false;
	}
	const closed = /^"(?:[^"]|"")*"$/.test(token.rawText);
	if (closed && offset > token.end - 1) { return false; }
	const before = previous(tokens, index)?.rawText.toLowerCase();
	if (before === '=') {
		return previous(tokens, index - 1)?.rawText.toLowerCase() === 'onaction';
	}
	if (before === ':=') {
		const parameter = previous(tokens, index - 1);
		return !!parameter && MACRO_PARAMETER.test(parameter.rawText);
	}
	return true;
}

/** The project procedures a macro-name string can name, `Module.Proc`, Declares and class modules left out. */
export function macroNameCandidates(ctx: SignatureHelpContext): MacroNameCandidate[] {
	const out = new Map<string, MacroNameCandidate>();
	for (const procedure of ctx.macroProcedures ?? ctx.projectProcedures ?? []) {
		const name = `${procedure.moduleName}.${procedure.name}`;
		if (!procedure.external && !out.has(name.toLowerCase())) {
			out.set(name.toLowerCase(), { name, procedure });
		}
	}
	return [...out.values()];
}

/**
 * The procedures to offer with the caret inside a procedure-name string, and
 * the text they replace: what is between the quotes. Undefined anywhere else,
 * the caret after a closing quote included.
 */
export function resolveMacroNameCompletions(
	source: string,
	offset: number,
	ctx: SignatureHelpContext = {},
): { contentSpan: Span; candidates: MacroNameCandidate[] } | undefined {
	const macro = macroNameStringAt(source, offset, ctx);
	if (!macro || offset > macro.contentSpan.end) {
		return undefined;
	}
	return { contentSpan: macro.contentSpan, candidates: macroNameCandidates(ctx) };
}

/** The procedure a macro-name string names: `Module.Proc`, or a bare `Proc` the project has once. */
export function macroNameTarget(text: string, ctx: SignatureHelpContext): VbaProcedureSignature | undefined {
	const trimmed = text.trim().replace(/^'[^']*'!/, '');
	const dot = trimmed.lastIndexOf('.');
	const moduleName = dot > 0 ? trimmed.slice(0, dot).toLowerCase() : undefined;
	const name = (dot > 0 ? trimmed.slice(dot + 1) : trimmed).toLowerCase();
	const matches = macroNameCandidates(ctx).filter((candidate) => candidate.procedure.name.toLowerCase() === name
		&& (moduleName === undefined || candidate.procedure.moduleName.toLowerCase() === moduleName));
	return matches.length === 1 ? matches[0].procedure : undefined;
}
