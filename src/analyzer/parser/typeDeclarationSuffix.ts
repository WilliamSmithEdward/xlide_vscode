export type TypeDeclarationSuffix = '$' | '%' | '&' | '!' | '#' | '@' | '^';

const TYPE_DECLARATION_SUFFIX_TYPES: Record<TypeDeclarationSuffix, string> = {
	$: 'String',
	'%': 'Integer',
	'&': 'Long',
	'!': 'Single',
	'#': 'Double',
	'@': 'Currency',
	'^': 'LongLong',
};

export function isTypeDeclarationSuffix(value: string | undefined): value is TypeDeclarationSuffix {
	return (
		value === '$' ||
		value === '%' ||
		value === '&' ||
		value === '!' ||
		value === '#' ||
		value === '@' ||
		value === '^'
	);
}

export function typeNameForDeclarationSuffix(
	suffix: string | undefined,
): string | undefined {
	return isTypeDeclarationSuffix(suffix) ? TYPE_DECLARATION_SUFFIX_TYPES[suffix] : undefined;
}
