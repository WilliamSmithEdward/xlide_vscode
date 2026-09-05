import { parseModule } from '../parser/parseModule';
import type { ModuleNode, ProcedureNode, Span } from '../parser/nodes';
import { detectEol } from '../../vbaSourceScan';
import { refactor, refuse, type VbaRefactorResult } from './refactorTypes';

/**
 * Implement Interface: a stub for every member an `Implements` promises and
 * the class has not written yet.
 *
 * Signatures are COPIED from the interface's own source text rather than
 * rebuilt from a symbol table. A rebuilt signature drifts - a missing
 * `Optional`, a lost `ParamArray`, a `ByRef` turned `ByVal` - and VBA rejects
 * an implementing member whose signature does not match to the letter, so the
 * copy is the only version that is always right.
 *
 * Bodies raise rather than return. A stub that silently returns the default
 * value is a bug that compiles, runs, and reports nothing.
 */

const NOT_IMPLEMENTED = "Err.Raise 5 'TODO: implement this interface member";

export interface ImplementInterfaceInput {
	/** The implementing class. */
	source: string;
	/** Which interface, when the class implements more than one. */
	interfaceName?: string;
	/** The source of every module the project has, keyed by module name. */
	moduleSources: Readonly<Record<string, string>>;
}

export function implementInterface(input: ImplementInterfaceInput): VbaRefactorResult {
	const module: ModuleNode = parseModule(input.source);
	const implemented = implementsNames(input.source);
	if (implemented.length === 0) {
		return refuse('This class implements no interface. Add an `Implements` statement first.');
	}

	const wanted = input.interfaceName
		?? (implemented.length === 1 ? implemented[0] : undefined);
	if (!wanted) {
		return refuse(`This class implements ${implemented.join(', ')}. Say which one to implement.`);
	}
	const name = implemented.find((n) => n.toLowerCase() === wanted.toLowerCase());
	if (!name) {
		return refuse(`This class does not implement '${wanted}'.`);
	}

	const interfaceSource = lookup(input.moduleSources, name);
	if (interfaceSource === undefined) {
		return refuse(`The project has no module called '${name}'.`);
	}

	const members = publicMembersOf(interfaceSource);
	if (members.length === 0) {
		return refuse(`'${name}' has no public members to implement.`);
	}

	const already = new Set(
		module.members
			.filter((member): member is ProcedureNode => member.kind === 'Procedure')
			.map((member) => member.name.toLowerCase()),
	);
	const missing = members.filter((member) => !already.has(`${name}_${member.name}`.toLowerCase()));
	if (missing.length === 0) {
		return refuse(`'${name}' is already implemented in full.`);
	}

	const eol = detectEol(input.source);
	const stubs = missing
		.map((member) => stubFor(name, member, eol))
		.join(eol + eol);
	const at = input.source.length;

	return refactor(
		`Implement ${missing.length} member${missing.length === 1 ? '' : 's'} of '${name}'`,
		[{
			span: { start: at, end: at },
			newText: (input.source.endsWith(eol) ? '' : eol) + eol + stubs + eol,
		}],
	);
}

/** A member the interface promises, with its header copied verbatim. */
interface InterfaceMember {
	name: string;
	/** `Property Get Total() As Long`, exactly as the interface writes it. */
	signature: string;
	/** The keyword that closes it: Sub, Function or Property. */
	closer: string;
	/** Whether a Property Get / Function returns an object, which needs `Set`. */
	isPropertyGet: boolean;
}

/**
 * The public members of an interface module. A public FIELD counts: VBA
 * exposes `Public Total As Long` on an interface as a Get/Let pair, and a
 * class that implements the interface has to write both.
 */
function publicMembersOf(source: string): InterfaceMember[] {
	const module = parseModule(source);
	const out: InterfaceMember[] = [];
	for (const member of module.members) {
		if (member.kind === 'Procedure') {
			if (/^private$/i.test(member.modifiers.find((m) => /^(public|private|friend)$/i.test(m)) ?? '')) {
				continue;
			}
			out.push({
				name: member.name,
				signature: headerText(source, member),
				closer: closerFor(member.procKind),
				isPropertyGet: member.procKind === 'PropertyGet',
			});
			continue;
		}
		if (member.kind === 'VariableGroup' && /^public$/i.test(member.modifier) && !member.isConst) {
			for (const decl of member.declarations) {
				const type = decl.asType ?? 'Variant';
				const isObject = !PRIMITIVES.has(type.toLowerCase());
				out.push({
					name: decl.name,
					signature: `Property Get ${decl.name}() As ${type}`,
					closer: 'Property',
					isPropertyGet: true,
				});
				out.push({
					name: decl.name,
					signature: `Property ${isObject ? 'Set' : 'Let'} ${decl.name}(ByVal RHS As ${type})`,
					closer: 'Property',
					isPropertyGet: false,
				});
			}
		}
	}
	return out;
}

const PRIMITIVES = new Set([
	'byte', 'boolean', 'integer', 'long', 'longlong', 'longptr', 'currency',
	'single', 'double', 'date', 'string', 'variant', 'decimal',
]);

/**
 * The member's header line as the interface wrote it, minus its access
 * modifier: an implementing member is always Private, and VBA rejects it
 * otherwise.
 */
function headerText(source: string, member: ProcedureNode): string {
	const line = source.slice(member.span.start, headerEnd(source, member.span));
	return line.trim().replace(/^\s*(?:Public|Private|Friend)\s+/i, '');
}

/** The end of the header, following any `_` line continuations. */
function headerEnd(source: string, span: Span): number {
	let at = source.indexOf('\n', span.start);
	if (at === -1) { return span.end; }
	while (/_[ \t]*\r?$/.test(source.slice(source.lastIndexOf('\n', at - 1) + 1, at))) {
		const next = source.indexOf('\n', at + 1);
		if (next === -1) { break; }
		at = next;
	}
	return Math.min(at, span.end);
}

function closerFor(procKind: ProcedureNode['procKind']): string {
	switch (procKind) {
		case 'Sub': return 'Sub';
		case 'Function': return 'Function';
		default: return 'Property';
	}
}

function stubFor(interfaceName: string, member: InterfaceMember, eol: string): string {
	// The name VBA requires: the interface, an underscore, the member.
	const renamed = member.signature.replace(
		new RegExp(`(\\b(?:Sub|Function|Property\\s+(?:Get|Let|Set))\\s+)${escapeForRegExp(member.name)}\\b`, 'i'),
		`$1${interfaceName}_${member.name}`,
	);
	return [
		`Private ${renamed}`,
		`    ${NOT_IMPLEMENTED}`,
		`End ${member.closer}`,
	].join(eol);
}

/** The interfaces an `Implements` line names, in source order. */
function implementsNames(source: string): string[] {
	const out: string[] = [];
	for (const match of source.matchAll(/^[ \t]*Implements[ \t]+([\p{L}_][\p{L}\p{M}\p{N}_.]*)/gimu)) {
		out.push(match[1]);
	}
	return out;
}

function lookup(sources: Readonly<Record<string, string>>, name: string): string | undefined {
	const lower = name.toLowerCase();
	const key = Object.keys(sources).find((k) => k.toLowerCase() === lower);
	return key === undefined ? undefined : sources[key];
}

function escapeForRegExp(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
