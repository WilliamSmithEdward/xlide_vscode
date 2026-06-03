export type VbaLimitBoundaryKind =
	| 'array-dimension-count'
	| 'argument-count'
	| 'continuation-count'
	| 'excel-worksheet-function-argument-count'
	| 'fixed-string-size'
	| 'identifier-length'
	| 'module-name-length'
	| 'physical-line-length'
	| 'string-literal-length';

export interface VbaLimitBoundaryFixture {
	id: string;
	kind: VbaLimitBoundaryKind;
	label: string;
	source: string;
	metadata: Readonly<Record<string, number | string>>;
}

const INDENT = '    ';

export function physicalLineLengthFixture(length: number): VbaLimitBoundaryFixture {
	assertPositiveInteger(length, 'length');
	const source = `'${'A'.repeat(length - 1)}`;
	return {
		id: `physical-line-length-${length}`,
		kind: 'physical-line-length',
		label: `Physical line length ${length}`,
		source,
		metadata: { lineLength: source.length },
	};
}

export function continuationCountFixture(count: number): VbaLimitBoundaryFixture {
	assertNonNegativeInteger(count, 'count');
	const lines = ['Sub T()'];
	if (count === 0) {
		lines.push(`${INDENT}Debug.Print 0`);
	} else {
		lines.push(`${INDENT}Debug.Print 0 _`);
		for (let i = 1; i <= count; i += 1) {
			lines.push(`${INDENT}+ ${i}${i < count ? ' _' : ''}`);
		}
	}
	lines.push('End Sub');
	return {
		id: `continuation-count-${count}`,
		kind: 'continuation-count',
		label: `Line continuation count ${count}`,
		source: lines.join('\n'),
		metadata: { continuationCount: count },
	};
}

export function stringLiteralLengthFixture(length: number): VbaLimitBoundaryFixture {
	assertNonNegativeInteger(length, 'length');
	const literal = 'A'.repeat(length);
	const source = ['Sub T()', `${INDENT}Dim value As String`, `${INDENT}value = "${literal}"`, 'End Sub'].join('\n');
	return {
		id: `string-literal-length-${length}`,
		kind: 'string-literal-length',
		label: `String literal length ${length}`,
		source,
		metadata: { literalLength: literal.length },
	};
}

export function fixedStringSizeFixture(size: number): VbaLimitBoundaryFixture {
	assertInteger(size, 'size');
	const source = ['Sub T()', `${INDENT}Dim buffer As String * ${size}`, 'End Sub'].join('\n');
	return {
		id: `fixed-string-size-${size}`,
		kind: 'fixed-string-size',
		label: `Fixed-length String size ${size}`,
		source,
		metadata: { fixedStringSize: size },
	};
}

export function identifierLengthFixture(length: number): VbaLimitBoundaryFixture {
	assertPositiveInteger(length, 'length');
	const identifier = `i${'d'.repeat(length - 1)}`;
	const source = ['Sub T()', `${INDENT}Dim ${identifier} As Long`, 'End Sub'].join('\n');
	return {
		id: `identifier-length-${length}`,
		kind: 'identifier-length',
		label: `Identifier length ${length}`,
		source,
		metadata: { identifier, identifierLength: identifier.length },
	};
}

export function moduleNameLengthFixture(length: number): VbaLimitBoundaryFixture {
	assertPositiveInteger(length, 'length');
	const moduleName = `M${'o'.repeat(length - 1)}`;
	return {
		id: `module-name-length-${length}`,
		kind: 'module-name-length',
		label: `Module name length ${length}`,
		source: 'Option Explicit\n',
		metadata: { moduleName, moduleNameLength: moduleName.length },
	};
}

export function argumentCountFixture(count: number): VbaLimitBoundaryFixture {
	assertNonNegativeInteger(count, 'count');
	const parameters = Array.from({ length: count }, (_, index) => `ByVal p${index + 1} As Long`);
	const argumentsText = Array.from({ length: count }, (_, index) => `${index + 1}`).join(', ');
	const call = argumentsText ? `${INDENT}Target ${argumentsText}` : `${INDENT}Target`;
	const signature = `Sub Target(${parameters.join(', ')})`;
	const source = ['Sub Caller()', call, 'End Sub', '', signature, 'End Sub'].join('\n');
	return {
		id: `argument-count-${count}`,
		kind: 'argument-count',
		label: `Argument count ${count}`,
		source,
		metadata: { argumentCount: count, parameterCount: parameters.length },
	};
}

export function arrayDimensionCountFixture(count: number): VbaLimitBoundaryFixture {
	assertPositiveInteger(count, 'count');
	const dimensions = Array.from({ length: count }, () => '1').join(', ');
	const source = ['Sub T()', `${INDENT}Dim values(${dimensions}) As Long`, 'End Sub'].join('\n');
	return {
		id: `array-dimension-count-${count}`,
		kind: 'array-dimension-count',
		label: `Array dimension count ${count}`,
		source,
		metadata: { dimensionCount: count },
	};
}

export function excelWorksheetFunctionArgumentCountFixture(count: number): VbaLimitBoundaryFixture {
	assertNonNegativeInteger(count, 'count');
	const argumentsText = Array.from({ length: count }, (_, index) => `${index + 1}`).join(', ');
	const source = [
		'Sub T()',
		`${INDENT}Dim value As Double`,
		`${INDENT}value = Application.WorksheetFunction.Sum(${argumentsText})`,
		'End Sub',
	].join('\n');
	return {
		id: `excel-worksheet-function-argument-count-${count}`,
		kind: 'excel-worksheet-function-argument-count',
		label: `Excel worksheet function argument count ${count}`,
		source,
		metadata: { argumentCount: count },
	};
}

export function allLimitBoundaryFixtures(): VbaLimitBoundaryFixture[] {
	return [
		continuationCountFixture(24),
		continuationCountFixture(25),
		physicalLineLengthFixture(1023),
		physicalLineLengthFixture(1024),
		stringLiteralLengthFixture(1023),
		stringLiteralLengthFixture(1024),
		fixedStringSizeFixture(65526),
		fixedStringSizeFixture(65527),
		identifierLengthFixture(255),
		identifierLengthFixture(256),
		moduleNameLengthFixture(31),
		moduleNameLengthFixture(32),
		argumentCountFixture(60),
		argumentCountFixture(61),
		arrayDimensionCountFixture(60),
		arrayDimensionCountFixture(61),
		excelWorksheetFunctionArgumentCountFixture(255),
		excelWorksheetFunctionArgumentCountFixture(256),
	];
}

function assertInteger(value: number, label: string): void {
	if (!Number.isInteger(value)) {
		throw new Error(`${label} must be an integer.`);
	}
}

function assertNonNegativeInteger(value: number, label: string): void {
	assertInteger(value, label);
	if (value < 0) {
		throw new Error(`${label} must be non-negative.`);
	}
}

function assertPositiveInteger(value: number, label: string): void {
	assertInteger(value, label);
	if (value <= 0) {
		throw new Error(`${label} must be positive.`);
	}
}
