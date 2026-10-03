/** Primitive types that refactoring assigns without Set. */
const NON_OBJECT_TYPES = new Set([
	'byte', 'boolean', 'integer', 'long', 'longlong', 'longptr', 'currency',
	'single', 'double', 'date', 'string', 'variant', 'decimal',
]);

/**
 * Refactoring assumes an unknown non-primitive type is a class. Without
 * project type resolution it cannot distinguish classes from Enums/UDTs;
 * preserve that existing assumption consistently in both property generators.
 */
export function isRefactorObjectType(declaredType: string): boolean {
	return !NON_OBJECT_TYPES.has(declaredType.toLowerCase());
}
