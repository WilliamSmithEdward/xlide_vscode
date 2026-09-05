import { parseModule } from '../parser/parseModule';
import type { ModuleNode, VariableDeclNode, VariableGroupNode } from '../parser/nodes';
import { detectEol } from '../../vbaSourceScan';
import { refactor, refuse, type VbaRefactorResult } from './refactorTypes';

/**
 * Encapsulate Field: a public module variable becomes private behind a
 * property pair that KEEPS ITS NAME, so not one call site is rewritten. The
 * backing field takes an `m_` prefix.
 *
 *     Public Total As Long
 *
 *     Private m_Total As Long
 *     Public Property Get Total() As Long
 *         Total = m_Total
 *     End Property
 *     Public Property Let Total(ByVal RHS As Long)
 *         m_Total = RHS
 *     End Property
 *
 * An object type gets `Property Set` instead of `Property Let`, and both it and
 * the getter assign with `Set` - the distinction VBA makes and the one that
 * decides whether the generated code compiles at all.
 */

/** Types that VBA assigns with `Set`. `Object` and `Variant` are not among them. */
const NON_OBJECT_TYPES = new Set([
	'byte', 'boolean', 'integer', 'long', 'longlong', 'longptr', 'currency',
	'single', 'double', 'date', 'string', 'variant', 'decimal',
]);

export interface EncapsulateFieldInput {
	source: string;
	/** Offset of the caret, anywhere inside the variable's declared name. */
	offset: number;
	/** Types the project declares as classes, so `As Widget` assigns with Set. */
	projectClassNames?: readonly string[];
}

export function encapsulateField(input: EncapsulateFieldInput): VbaRefactorResult {
	const module: ModuleNode = parseModule(input.source);
	const found = fieldAt(module, input.offset);
	if (!found) {
		return refuse('Put the caret on a module-level variable to encapsulate it.');
	}
	const { group, decl } = found;

	if (group.isConst) {
		return refuse(`'${decl.name}' is a Const, which has no value to set.`);
	}
	if (/^private$/i.test(group.modifier)) {
		return refuse(`'${decl.name}' is already Private, so nothing outside the module reads it.`);
	}
	if (group.withEvents) {
		return refuse(`'${decl.name}' is declared WithEvents, and a property cannot raise its events.`);
	}
	if (decl.isArray) {
		return refuse(`'${decl.name}' is an array, and VBA properties cannot return one by reference.`);
	}
	if (group.declarations.length > 1) {
		return refuse(
			`'${decl.name}' shares its declaration with `
			+ `${group.declarations.filter((d) => d !== decl).map((d) => `'${d.name}'`).join(', ')}. `
			+ 'Split the declaration first.',
		);
	}
	if (declarationSpansLines(input.source, group)) {
		return refuse(`The declaration of '${decl.name}' is continued across lines. Join it first.`);
	}

	const backing = `m_${decl.name}`;
	const taken = takenNames(module);
	if (taken.has(backing.toLowerCase())) {
		return refuse(`The module already has something called '${backing}'.`);
	}
	const propertyName = decl.name;
	for (const member of module.members) {
		if (member.kind === 'Procedure' && member.name.toLowerCase() === propertyName.toLowerCase()) {
			return refuse(`The module already has a procedure called '${propertyName}'.`);
		}
	}

	const eol = detectEol(input.source);
	const declaredType = typeOf(decl);
	const isObject = isObjectType(declaredType, input.projectClassNames);
	const set = isObject ? 'Set ' : '';
	// `RHS` is the VBE's own name for a property's value parameter, which is
	// what a reader of generated VBA expects to see.
	const lines = [
		`Private ${backing}${asClause(decl)}`,
		'',
		`Public Property Get ${propertyName}()${returnClause(declaredType)}`,
		`    ${set}${propertyName} = ${backing}`,
		'End Property',
		'',
		`Public Property ${isObject ? 'Set' : 'Let'} ${propertyName}(ByVal RHS${returnClause(declaredType)})`,
		`    ${set}${backing} = RHS`,
		'End Property',
	];

	return refactor(
		`Encapsulate '${decl.name}' behind a property`,
		[{ span: group.span, newText: lines.join(eol) }],
	);
}

function fieldAt(
	module: ModuleNode,
	offset: number,
): { group: VariableGroupNode; decl: VariableDeclNode } | undefined {
	for (const member of module.members) {
		if (member.kind !== 'VariableGroup') {
			continue;
		}
		for (const decl of member.declarations) {
			const span = decl.nameSpan ?? decl.span;
			if (offset >= span.start && offset <= span.end) {
				return { group: member, decl };
			}
		}
		// The caret anywhere on a single-name declaration means that name.
		if (member.declarations.length === 1
			&& offset >= member.span.start && offset <= member.span.end) {
			return { group: member, decl: member.declarations[0] };
		}
	}
	return undefined;
}

/** Every module-level name, so a generated one cannot collide with it. */
function takenNames(module: ModuleNode): Set<string> {
	const names = new Set<string>();
	for (const member of module.members) {
		if (member.kind === 'VariableGroup') {
			for (const decl of member.declarations) { names.add(decl.name.toLowerCase()); }
		} else if (member.kind === 'Procedure') {
			names.add(member.name.toLowerCase());
		} else if (member.kind === 'Enum' || member.kind === 'Type' || member.kind === 'Declare') {
			names.add(member.name.toLowerCase());
		}
	}
	return names;
}

function declarationSpansLines(source: string, group: VariableGroupNode): boolean {
	return /[\r\n]/.test(source.slice(group.span.start, group.span.end));
}

/**
 * `Variant` is what VBA gives a name nothing narrows, so it is the honest
 * answer for an untyped field rather than a failure. The parser resolves a
 * type suffix into `asType` for us, so `Count%` reads as Integer here.
 */
function typeOf(decl: VariableDeclNode): string {
	return decl.asType ?? 'Variant';
}

/**
 * The backing field's `As` clause. A suffix is written out in full, which is
 * the one deliberate change to the declaration: `Private m_Count As Integer`
 * says the same thing as `Private m_Count%` and matches the property beside it.
 */
function asClause(decl: VariableDeclNode): string {
	if (!decl.asType) {
		return '';
	}
	return ` As ${decl.isNew ? 'New ' : ''}${decl.asType}${decl.fixedLength ? ` * ${decl.fixedLength}` : ''}`;
}

function returnClause(declaredType: string): string {
	return ` As ${declaredType}`;
}

function isObjectType(declaredType: string, projectClassNames: readonly string[] = []): boolean {
	const lower = declaredType.toLowerCase();
	if (NON_OBJECT_TYPES.has(lower)) {
		return false;
	}
	if (lower === 'object' || lower.includes('.')) {
		return true;
	}
	if (projectClassNames.some((name) => name.toLowerCase() === lower)) {
		return true;
	}
	// An unknown name is a type the module names but this call cannot see -
	// a class, an Enum, a UDT. Enums and UDTs are Let; a class is Set. Without
	// the project there is no way to tell, so the safe read is the one that
	// still compiles for the common case: a bare unknown name is a class.
	return !/^(?:byte|boolean|integer|long|longlong|longptr|currency|single|double|date|string|variant|decimal)$/i
		.test(lower);
}
