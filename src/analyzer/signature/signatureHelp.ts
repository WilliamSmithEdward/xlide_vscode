// Signature help (parameter info) resolver - the VBE "Quick Info" / call tip.
//
// Pure analyzer code: given a source string and a caret offset, it determines
// whether the caret is inside a call's argument list and, if so, returns the
// callee's signature plus the index of the active parameter. No vscode or I/O
// dependencies, so the resolver is unit-testable.
//
// Three signature sources are consulted, in order of specificity:
//   1. Host members  (e.g. Workbooks.Open) - verified Office object-model
//      signatures from excelObjectModel.ts (resolveHostMemberSignature).
//   2. User procedures (Sub/Function/Property) in the current module - built
//      from the parsed AST so ByVal/Optional/ParamArray detail is exact.
//   3. Built-in runtime functions (MsgBox, Left, ...) - the verified signature
//      strings in vbaRuntime.ts.
//
// Both parenthesized calls -- Workbooks.Open(... -- and parenless call
// statements -- Workbooks.Open "file" -- are recognised, matching the VBE which
// shows the call tip in both forms. Signatures are never invented: a callee
// with no known signature simply yields no tip.

import { tokenize } from '../lexer/tokenize';
import { VbaToken } from '../lexer/tokenKinds';
import { parseModule } from '../parser/parseModule';
import { ParameterNode, ProcedureNode } from '../parser/nodes';
import { resolveReceiverTypeAt, MemberCompletionContext } from '../completion/memberAccess';
import { resolveHostMemberSignature } from '../host/hostModel';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';

/** A single parameter slot within a signature label. */
export interface SignatureParameter {
	/** The parameter text exactly as it appears in `SignatureInfo.label`. */
	label: string;
}

/** A resolved call tip: the signature line and which parameter is active. */
export interface SignatureInfo {
	/** Full signature line, e.g. `Open(Filename As String, [ReadOnly]) As Workbook`. */
	label: string;
	/** Parameters in declaration order. */
	parameters: SignatureParameter[];
	/** Zero-based index of the active parameter (clamped to the last param). */
	activeParameter: number;
}

/** Context for signature resolution. */
export interface SignatureHelpContext extends MemberCompletionContext {
	/**
	 * Source of the module that owns user procedures to match bare callees
	 * against. Defaults to the analysed `source` when omitted (the common case
	 * where the caret and the procedures live in the same module).
	 */
	moduleSource?: string;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

/** Statement-leading keywords that must never be treated as parenless calls. */
export const STATEMENT_KEYWORDS = new Set([
	'dim', 'set', 'let', 'const', 'redim', 'static', 'global', 'public',
	'private', 'friend', 'if', 'elseif', 'else', 'then', 'for', 'next', 'do',
	'loop', 'while', 'wend', 'with', 'select', 'case', 'end', 'exit', 'on',
	'goto', 'gosub', 'return', 'declare', 'type', 'enum', 'property', 'sub',
	'function', 'option', 'get', 'resume', 'error', 'stop', 'open', 'close',
	'print', 'write', 'input', 'line', 'name', 'kill', 'erase', 'lock',
	'unlock', 'seek', 'put', 'mkdir', 'rmdir', 'chdir', 'chdrive', 'load',
	'unload',
]);

function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

/** A located call whose argument list contains the caret. */
interface CallSite {
	calleeName: string;
	/** True when the callee is `receiver.Member`. */
	isMember: boolean;
	/** Absolute offset just past the callee identifier. */
	calleeEndOffset: number;
	/** Zero-based index of the argument the caret is in. */
	activeParameter: number;
}

/**
 * Locates the innermost enclosing *parenthesized* call whose argument list the
 * caret sits in. Returns undefined when the caret is not inside any call paren
 * (a grouping paren or no paren).
 */
function findParenCall(tokens: VbaToken[]): CallSite | undefined {
	interface Frame {
		isCall: boolean;
		openIndex: number;
		commaCount: number;
	}
	const stack: Frame[] = [];
	for (let k = 0; k < tokens.length; k += 1) {
		const t = tokens[k];
		if (t.kind === 'newline' || t.rawText === ':') {
			stack.length = 0;
			continue;
		}
		const r = t.rawText;
		if (r === '(') {
			const prev = tokens[k - 1];
			// A call paren is one preceded directly by an identifier-like token.
			// A paren after ')' (default-member indexing) or an operator is a
			// grouping/index paren and offers no signature.
			const isCall = !!prev && isIdentLike(prev);
			stack.push({ isCall, openIndex: k, commaCount: 0 });
		} else if (r === ')') {
			stack.pop();
		} else if (r === ',' && stack.length > 0) {
			stack[stack.length - 1].commaCount += 1;
		}
	}
	for (let i = stack.length - 1; i >= 0; i -= 1) {
		if (stack[i].isCall) {
			const open = stack[i].openIndex;
			const callee = tokens[open - 1];
			const isMember = open - 2 >= 0 && tokens[open - 2].rawText === '.';
			return {
				calleeName: callee.rawText,
				isMember,
				calleeEndOffset: callee.end,
				activeParameter: stack[i].commaCount,
			};
		}
	}
	return undefined;
}

/**
 * Locates a parenless call statement -- e.g. `Workbooks.Open "file"` or
 * `MyMacro a, b` -- whose argument region the caret sits in. Conservative: only
 * fires when the statement begins with a non-keyword identifier chain followed
 * by whitespace (the argument separator) and contains no top-level `=` (which
 * would make it an assignment, not a call).
 */
function findParenlessCall(
	tokens: VbaToken[],
	source: string,
	offset: number,
): CallSite | undefined {
	let start = 0;
	for (let k = tokens.length - 1; k >= 0; k -= 1) {
		if (tokens[k].kind === 'newline' || tokens[k].rawText === ':') {
			start = k + 1;
			break;
		}
	}
	const stmt = tokens.slice(start);
	if (stmt.length === 0) {
		return undefined;
	}
	let idx = 0;
	if (stmt[0].rawText.toLowerCase() === 'call') {
		idx = 1;
	}
	if (idx >= stmt.length || !isIdentLike(stmt[idx])) {
		return undefined;
	}
	if (STATEMENT_KEYWORDS.has(stmt[idx].rawText.toLowerCase())) {
		return undefined;
	}
	// Collect the dotted callee chain.
	let j = idx;
	for (;;) {
		if (!isIdentLike(stmt[j])) {
			return undefined;
		}
		if (j + 1 < stmt.length && stmt[j + 1].rawText === '.') {
			j += 2;
			continue;
		}
		break;
	}
	const callee = stmt[j];
	const isMember = j - 1 >= idx && stmt[j - 1].rawText === '.';
	const afterTokens = stmt.slice(j + 1);
	const gap = source.slice(callee.end, offset);
	const argsStarted = afterTokens.length > 0 || /\s/.test(gap);
	if (!argsStarted) {
		return undefined;
	}
	let depth = 0;
	let commaCount = 0;
	for (const t of afterTokens) {
		const r = t.rawText;
		if (r === '(' || r === '[') {
			depth += 1;
		} else if (r === ')' || r === ']') {
			depth -= 1;
		} else if (depth === 0) {
			if (r === '=') {
				return undefined; // assignment, not a call statement
			}
			if (r === ',') {
				commaCount += 1;
			}
		}
	}
	return {
		calleeName: callee.rawText,
		isMember,
		calleeEndOffset: callee.end,
		activeParameter: commaCount,
	};
}

function formatParam(p: ParameterNode): string {
	// Matches the VBE call-tip convention: optional parameters are shown in
	// [brackets] (the "Optional" keyword and ByVal/ByRef are omitted), while
	// ParamArray, the type, and any default value are kept.
	let s = p.name;
	if (p.isArray) {
		s += '()';
	}
	if (p.asType) {
		s += ` As ${p.asType}`;
	}
	if (p.defaultRaw) {
		s += ` = ${p.defaultRaw}`;
	}
	if (p.paramArray) {
		s = `ParamArray ${s}`;
	}
	return p.optional ? `[${s}]` : s;
}

function userProcSignature(proc: ProcedureNode): string {
	// No leading Sub/Function keyword, matching the runtime/host signature style
	// (e.g. "Left(String, Length) As String"); the trailing "As <type>" marks a
	// Function/Property Get return.
	const params = proc.params.map(formatParam).join(', ');
	const ret = proc.returnType ? ` As ${proc.returnType}` : '';
	return `${proc.name}(${params})${ret}`;
}

function findUserProc(source: string, name: string): ProcedureNode | undefined {
	const lower = name.toLowerCase();
	const module = parseModule(source);
	for (const member of module.members) {
		if (member.kind === 'Procedure' && member.name.toLowerCase() === lower) {
			return member;
		}
	}
	return undefined;
}

/** Splits a parameter list on top-level commas, respecting (), [] and strings. */
function splitTopLevel(text: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let inStr = false;
	let startPos = 0;
	for (let i = 0; i < text.length; i += 1) {
		const c = text[i];
		if (inStr) {
			if (c === '"') {
				inStr = false;
			}
			continue;
		}
		if (c === '"') {
			inStr = true;
		} else if (c === '(' || c === '[') {
			depth += 1;
		} else if (c === ')' || c === ']') {
			depth -= 1;
		} else if (c === ',' && depth === 0) {
			out.push(text.slice(startPos, i));
			startPos = i + 1;
		}
	}
	out.push(text.slice(startPos));
	return out;
}

/** Parses a signature string into its label and ordered parameter substrings. */
function parseSignature(sig: string): { label: string; params: string[] } {
	const open = sig.indexOf('(');
	if (open < 0) {
		return { label: sig, params: [] };
	}
	let depth = 0;
	let close = -1;
	for (let i = open; i < sig.length; i += 1) {
		const c = sig[i];
		if (c === '(') {
			depth += 1;
		} else if (c === ')') {
			depth -= 1;
			if (depth === 0) {
				close = i;
				break;
			}
		}
	}
	if (close < 0) {
		return { label: sig, params: [] };
	}
	const inner = sig.slice(open + 1, close);
	const params = splitTopLevel(inner)
		.map((s) => s.trim())
		.filter((s) => s.length > 0);
	return { label: sig, params };
}

function signatureForCallee(
	site: CallSite,
	source: string,
	ctx: SignatureHelpContext,
): string | undefined {
	if (site.isMember) {
		const recvType = resolveReceiverTypeAt(source, site.calleeEndOffset, ctx);
		if (!recvType) {
			return undefined;
		}
		return resolveHostMemberSignature(recvType, site.calleeName, ctx.model);
	}
	const proc = findUserProc(ctx.moduleSource ?? source, site.calleeName);
	if (proc) {
		return userProcSignature(proc);
	}
	return resolveRuntimeFunction(site.calleeName)?.signature;
}

/**
 * Resolves the call tip active at `offset`, or undefined when the caret is not
 * inside a recognised call's argument list or the callee has no known
 * signature. Never throws: any internal failure yields undefined so the editor
 * is never disrupted.
 */
export function resolveSignatureHelp(
	source: string,
	offset: number,
	ctx: SignatureHelpContext = {},
): SignatureInfo | undefined {
	try {
		if (offset < 0) {
			return undefined;
		}
		const prefix = source.slice(0, offset);
		const tokens = tokenize(prefix).filter((t) => t.kind !== 'comment');
		if (tokens.length === 0) {
			return undefined;
		}
		const site = findParenCall(tokens) ?? findParenlessCall(tokens, source, offset);
		if (!site) {
			return undefined;
		}
		const sig = signatureForCallee(site, source, ctx);
		if (!sig) {
			return undefined;
		}
		const parsed = parseSignature(sig);
		let active = site.activeParameter;
		if (parsed.params.length === 0) {
			active = 0;
		} else if (active > parsed.params.length - 1) {
			active = parsed.params.length - 1;
		}
		if (active < 0) {
			active = 0;
		}
		return {
			label: parsed.label,
			parameters: parsed.params.map((p) => ({ label: p })),
			activeParameter: active,
		};
	} catch {
		return undefined;
	}
}
