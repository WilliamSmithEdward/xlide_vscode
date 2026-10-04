import { parseModule } from '../parser/parseModule';
import { blockHeaderLineSpan } from '../parser/physicalLineSpans';
import type { ModuleNode, ProcedureNode, ProcKind } from '../parser/nodes';
import { detectEol } from '../../vbaSourceScan';
import { refactor, refuse, type VbaRefactorResult } from './refactorTypes';
import { escapeForRegExp, lookupModuleSource } from './shared';
import { isRefactorObjectType } from './typeKinds';

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

	const interfaceSource = lookupModuleSource(input.moduleSources, name);
	if (interfaceSource === undefined) {
		return refuse(`The project has no module called '${name}'.`);
	}

	const members = publicMembersOf(interfaceSource);
	if (members.length === 0) {
		return refuse(`'${name}' has no public members to implement.`);
	}

	// The refusal checks above do not need the implementing class AST.
	const module: ModuleNode = parseModule(input.source);

	const requiredNames = new Set(members.map(member => `${name}_${member.name}`.toLowerCase()));
	const kindsByName = new Map<string, Set<ProcKind>>();
	for (const member of module.members) {
		if (member.kind !== 'Procedure') { continue; }
		const key = member.name.toLowerCase();
		if (!requiredNames.has(key)) { continue; }
		let kinds = kindsByName.get(key);
		if (!kinds) { kinds = new Set(); kindsByName.set(key, kinds); }
		kinds.add(member.procKind);
	}
	const missing: InterfaceMember[] = [];
	for (const member of members) {
		const qualifiedName = `${name}_${member.name}`;
		const kinds = kindsByName.get(qualifiedName.toLowerCase());
		if (kinds?.has(member.procKind)) { continue; }
		// Property Get/Let/Set may share a name; Sub/Function cannot share one
		// with a different callable kind. Refuse instead of creating a collision.
		if (kinds && (member.procKind === 'Sub' || member.procKind === 'Function'
			|| kinds.has('Sub') || kinds.has('Function'))) {
			return refuse(`The class already has '${qualifiedName}' as a different procedure kind. Rename or correct it before implementing '${name}'.`);
		}
		missing.push(member);
	}
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
	procKind: ProcKind;
	/** `Property Get Total() As Long`, exactly as the interface writes it. */
	signature: string;
	/** The keyword that closes it: Sub, Function or Property. */
	closer: string;
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
				procKind: member.procKind,
				signature: headerText(source, member),
				closer: closerFor(member.procKind),
			});
			continue;
		}
		if (member.kind === 'VariableGroup' && /^public$/i.test(member.modifier) && !member.isConst) {
			for (const decl of member.declarations) {
				const type = decl.asType ?? 'Variant';
				const isObject = isRefactorObjectType(type);
				out.push({
					name: decl.name,
					procKind: 'PropertyGet',
					signature: `Property Get ${decl.name}() As ${type}`,
					closer: 'Property',
				});
				out.push({
					name: decl.name,
					procKind: isObject ? 'PropertySet' : 'PropertyLet',
					signature: `Property ${isObject ? 'Set' : 'Let'} ${decl.name}(ByVal RHS As ${type})`,
					closer: 'Property',
				});
			}
		}
	}
	return out;
}

/**
 * The member's header line as the interface wrote it, minus its access
 * modifier: an implementing member is always Private, and VBA rejects it
 * otherwise.
 */
function headerText(source: string, member: ProcedureNode): string {
	const header = blockHeaderLineSpan(source, member.span);
	const line = source.slice(header.start, header.end);
	return line.trim().replace(/^\s*(?:Public|Private|Friend)\s+/i, '');
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
export function implementsNames(source: string): string[] {
	const out: string[] = [];
	for (const match of source.matchAll(/^[ \t]*Implements[ \t]+([\p{L}_][\p{L}\p{M}\p{N}_.]*)/gimu)) {
		out.push(match[1]);
	}
	return out;
}
