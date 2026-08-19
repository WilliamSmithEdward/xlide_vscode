// Signature help (parameter info) resolver - the VBE "Quick Info" / call tip.
//
// Pure analyzer code: given a source string and a caret offset, it determines
// whether the caret is inside a call's argument list and, if so, returns the
// callee's signature plus the index of the active parameter. No vscode or I/O
// dependencies, so the resolver is unit-testable.
//
// Three signature sources are consulted, in order of specificity:
//   1. Host members  (e.g. Workbooks.Open) - verified Office object-model
//      signatures from generated HostMember metadata or excelObjectModel.ts
//      fallbacks (via member completion / resolveHostMemberSignature).
//   2. User procedures and external Declares in the current module - built
//      from the parsed AST so ByVal/Optional/ParamArray detail is exact.
//   3. Built-in runtime functions (MsgBox, Left, ...) - the verified signature
//      strings in vbaRuntime.ts.
//
// Both parenthesized calls -- Workbooks.Open(... -- and parenless call
// statements -- Workbooks.Open "file" -- are recognised, matching the VBE which
// shows the call tip in both forms. Signatures are never invented: a callee
// with no known signature simply yields no tip.

import { parseModule } from '../parser/parseModule';
import { DeclareNode, ParameterNode, ProcedureNode } from '../parser/nodes';
import { resolveMemberCompletionNamed, MemberCompletionContext } from '../completion/memberAccess';
import { getHostType, resolveHostGlobalMember } from '../host/hostModel';
import { resolveRuntimeFunction, runtimeAllowsExplicitCall } from '../runtime/vbaRuntime';
import { extractLeadingDoc } from '../docs/docComment';
import { DocRegistry } from '../docs/docRegistry';
import { VbaDoc, hasDocContent, renderParamDocMarkdown, renderSignatureDocMarkdown } from '../docs/docModel';
import {
	procedureDeclarationSignature,
	type VbaProcedureSignature,
	type VbaProcedureParam,
	procedureSignatureLabel,
} from '../symbols/symbolModel';
import {
	findActiveCallSite,
	type VbaCallSite as CallSite,
	STATEMENT_KEYWORDS,
} from '../call/callContext';

/** A single parameter slot within a signature label. */
export interface SignatureParameter {
	/** The parameter text exactly as it appears in `SignatureInfo.label`. */
	label: string;
	/** Markdown note for this parameter (from `<param name="...">`). */
	documentation?: string;
}

/** A resolved call tip: the signature line and which parameter is active. */
export interface SignatureInfo {
	/** Full signature line, e.g. `Open(Filename As String, [ReadOnly]) As Workbook`. */
	label: string;
	/** Parameters in declaration order. */
	parameters: SignatureParameter[];
	/** Zero-based index of the active parameter (clamped to the last param). */
	activeParameter: number;
	/** Markdown summary (summary/returns/remarks) for the whole signature. */
	documentation?: string;
	/** Plain detail lines for metadata such as external Declare Lib/Alias. */
	details?: string[];
}

/** Context for signature resolution. */
export interface SignatureHelpContext extends MemberCompletionContext {
	/** Name of the module being edited, used for Declare call-tip detail text. */
	moduleName?: string;
	/**
	 * Source of the module that owns user procedures to match bare callees
	 * against. Defaults to the analysed `source` when omitted (the common case
	 * where the caret and the procedures live in the same module).
	 */
	moduleSource?: string;
	/** Exported project procedures/Declares visible as bare calls from this module. */
	projectProcedures?: readonly VbaProcedureSignature[];
	/** Developer-defined external documentation (overrides the curated library). */
	docRegistry?: DocRegistry;
}

export { STATEMENT_KEYWORDS };

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

function declareParam(param: ParameterNode): VbaProcedureParam {
	const out: VbaProcedureParam = {
		name: param.name,
		type: param.asType,
		optional: param.optional,
		paramArray: param.paramArray,
		isArray: param.isArray,
		defaultRaw: param.defaultRaw,
	};
	if (param.byVal) {
		out.byVal = true;
	} else if (param.byRef) {
		out.byRef = true;
	}
	return out;
}

function userDeclareSignature(declare: DeclareNode): string {
	return procedureDeclarationSignature({
		name: declare.name,
		kind: declare.isFunction ? 'function' : 'sub',
		params: declare.params.map(declareParam),
		returnType: declare.returnType,
		external: true,
		ptrSafe: declare.ptrSafe,
		libName: declare.libName,
		aliasName: declare.aliasName,
	});
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

function findUserDeclare(source: string, name: string): DeclareNode | undefined {
	const lower = name.toLowerCase();
	const module = parseModule(source);
	for (const member of module.members) {
		if (member.kind === 'Declare' && member.name.toLowerCase() === lower) {
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
		return memberCompletionForCallee(site, source, ctx)?.signature;
	}
	const proc = findUserProc(ctx.moduleSource ?? source, site.calleeName);
	if (proc) {
		return userProcSignature(proc);
	}
	const declare = findUserDeclare(ctx.moduleSource ?? source, site.calleeName);
	if (declare) {
		return userDeclareSignature(declare);
	}
	const projectProc = findProjectProcedure(ctx.projectProcedures, site.calleeName);
	if (projectProc) {
		return projectProc.signature ?? (
			projectProc.external
				? procedureDeclarationSignature(projectProc)
				: procedureSignatureLabel(projectProc)
		);
	}
	// A bare call may bind to the host's hidden Global interface - Word's
	// InchesToPoints, Excel's Union (issue #34). The VBA runtime's own names
	// keep winning first.
	return resolveRuntimeFunction(site.calleeName)?.signature
		?? resolveHostGlobalMember(site.calleeName, ctx.model)?.signature;
}

function memberCompletionForCallee(
	site: CallSite,
	source: string,
	ctx: SignatureHelpContext,
) {
	return resolveMemberCompletionNamed(source, site.calleeEndOffset, site.calleeName, ctx);
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
		const site = findActiveCallSite(source, offset);
		if (!site) {
			return undefined;
		}
		if (isForbiddenExplicitRuntimeCall(site)) {
			return undefined;
		}
		const doc = docForCallee(site, source, ctx);
		const details = signatureDetailsForCallee(site, source, ctx);
		// External metadata can supply a signature for a callee the curated
		// library cannot resolve (developer-defined overrides the library).
		const sig = signatureForCallee(site, source, ctx) ?? doc?.signature;
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
			parameters: parsed.params.map((p) => ({
				label: p,
				documentation: paramDocFor(doc, p),
			})),
			activeParameter: active,
			documentation: signatureDocumentation(doc, details),
			details: details.length > 0 ? details : undefined,
		};
	} catch {
		return undefined;
	}
}

function signatureDocumentation(
	doc: VbaDoc | undefined,
	details: readonly string[],
): string | undefined {
	const renderedDoc = hasDocContent(doc)
		? renderSignatureDocMarkdown(doc) || undefined
		: undefined;
	if (details.length === 0) {
		return renderedDoc;
	}
	const renderedDetails = details.join('  \n');
	return renderedDoc
		? `${renderedDoc}\n\n${renderedDetails}`
		: renderedDetails;
}

function signatureDetailsForCallee(
	site: CallSite,
	source: string,
	ctx: SignatureHelpContext,
): string[] {
	if (site.isMember) {
		return [];
	}
	const moduleSource = ctx.moduleSource ?? source;
	const declare = findUserDeclare(moduleSource, site.calleeName);
	if (declare) {
		return declareDetails({
			moduleName: ctx.moduleName,
			visibility: declare.visibility,
			libName: declare.libName,
			aliasName: declare.aliasName,
		});
	}
	const projectProc = findProjectProcedure(ctx.projectProcedures, site.calleeName);
	if (!projectProc?.external) {
		return [];
	}
	return declareDetails({
		moduleName: projectProc.moduleName,
		visibility: projectProc.visibility,
		libName: projectProc.libName,
		aliasName: projectProc.aliasName,
	});
}

function declareDetails(info: {
	moduleName?: string;
	visibility?: string;
	libName?: string;
	aliasName?: string;
}): string[] {
	const details: string[] = ['External declaration'];
	if (info.moduleName) {
		details.push(`Declared in Module: ${info.moduleName}`);
	}
	details.push(`Visibility: ${info.visibility ?? 'Public'}`);
	if (info.libName) {
		details.push(`Lib: ${info.libName}`);
	}
	if (info.aliasName) {
		details.push(`Alias: ${info.aliasName}`);
	}
	return details;
}

function isForbiddenExplicitRuntimeCall(site: CallSite): boolean {
	if (!site.isExplicitCall || site.isMember) {
		return false;
	}
	const runtime = resolveRuntimeFunction(site.calleeName);
	return !!runtime && !runtimeAllowsExplicitCall(runtime);
}

/**
 * Resolves developer-defined documentation for the call site: an inline `'''`
 * comment on a user procedure wins, then an external metadata entry. Host and
 * runtime callees are documented through the registry only.
 */
function docForCallee(
	site: CallSite,
	source: string,
	ctx: SignatureHelpContext,
): VbaDoc | undefined {
	if (site.isMember) {
		const member = memberCompletionForCallee(site, source, ctx);
		const external = member
			? externalDocForMember(ctx, site.calleeName, member.owner)
			: ctx.docRegistry?.lookup(site.calleeName);
		if (!member) {
			return external;
		}
		const hostMember = !!getHostType(member.owner, ctx.model);
		return hostMember ? external ?? member.doc : member.doc ?? external;
	}
	if (!site.isMember) {
		const moduleSource = ctx.moduleSource ?? source;
		const proc = findUserProc(moduleSource, site.calleeName);
		if (proc) {
			const inline = extractLeadingDoc(moduleSource, proc.span.start);
			if (inline) {
				return inline;
			}
		}
		const declare = findUserDeclare(moduleSource, site.calleeName);
		if (declare) {
			const inline = extractLeadingDoc(moduleSource, declare.span.start);
			if (inline) {
				return inline;
			}
		}
		const projectProc = findProjectProcedure(ctx.projectProcedures, site.calleeName);
		if (projectProc) {
			return projectProc.doc ?? ctx.docRegistry?.lookup(site.calleeName, projectProc.moduleName);
		}
	}
	return ctx.docRegistry?.lookup(site.calleeName);
}

function findProjectProcedure(
	procedures: readonly VbaProcedureSignature[] | undefined,
	name: string,
): VbaProcedureSignature | undefined {
	const matches = (procedures ?? []).filter(
		(procedure) => procedure.name.toLowerCase() === name.toLowerCase(),
	);
	return matches.length === 1 ? matches[0] : undefined;
}

function externalDocForMember(
	ctx: SignatureHelpContext,
	name: string,
	owner: string,
): VbaDoc | undefined {
	const qualifier = getHostType(owner, ctx.model)?.displayName ?? displayTypeName(owner);
	return ctx.docRegistry?.lookup(name, qualifier) ?? ctx.docRegistry?.lookup(name);
}

function displayTypeName(typeName: string): string {
	const dot = typeName.lastIndexOf('.');
	return dot >= 0 ? typeName.slice(dot + 1) : typeName;
}

/** Returns the `<param>` note matching a signature parameter substring, if any. */
function paramDocFor(doc: VbaDoc | undefined, paramLabel: string): string | undefined {
	if (!doc || doc.params.length === 0) {
		return undefined;
	}
	const name = leadingIdentifier(paramLabel);
	if (!name) {
		return undefined;
	}
	const lower = name.toLowerCase();
	const match = doc.params.find((p) => p.name.toLowerCase() === lower);
	return match ? renderParamDocMarkdown(match) : undefined;
}

/** Extracts the parameter name from a signature parameter slot text. */
function leadingIdentifier(paramLabel: string): string | undefined {
	const modifiers = new Set(['optional', 'byval', 'byref', 'paramarray']);
	// A parameter can be named in any script, so this cannot be ASCII-only:
	// `ByVal значение As String` would otherwise yield no name at all.
	const re = /[\p{L}_][\p{L}\p{M}\p{N}_]*/gu;
	let m: RegExpExecArray | null;
	while ((m = re.exec(paramLabel)) !== null) {
		if (!modifiers.has(m[0].toLowerCase())) {
			return m[0];
		}
	}
	return undefined;
}
