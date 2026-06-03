import { describe, expect, it } from 'vitest';
import {
	allLimitBoundaryFixtures,
	arrayDimensionCountFixture,
	argumentCountFixture,
	continuationCountFixture,
	excelWorksheetFunctionArgumentCountFixture,
	fixedStringSizeFixture,
	identifierLengthFixture,
	moduleNameLengthFixture,
	physicalLineLengthFixture,
	stringLiteralLengthFixture,
} from './helpers/vbaLimitBoundaryFixtures';

describe('VBA limit boundary fixture builders', () => {
	it('creates a deterministic, unique fixture surface for corpus promotion', () => {
		const first = allLimitBoundaryFixtures();
		const second = allLimitBoundaryFixtures();

		expect(first.map((fixture) => fixture.id)).toEqual(second.map((fixture) => fixture.id));
		expect(first.map((fixture) => fixture.source)).toEqual(second.map((fixture) => fixture.source));
		expect(new Set(first.map((fixture) => fixture.id)).size).toBe(first.length);
		expect(new Set(first.map((fixture) => fixture.kind))).toEqual(new Set([
			'array-dimension-count',
			'argument-count',
			'continuation-count',
			'excel-worksheet-function-argument-count',
			'fixed-string-size',
			'identifier-length',
			'module-name-length',
			'physical-line-length',
			'string-literal-length',
		]));
	});

	it('generates exact continuation-count fixtures without relying on hand-written bodies', () => {
		const fixture = continuationCountFixture(25);

		expect(fixture.metadata.continuationCount).toBe(25);
		expect(fixture.source.match(/ _$/gm)).toHaveLength(25);
	});

	it('generates exact physical-line and string-literal length fixtures', () => {
		const physicalLine = physicalLineLengthFixture(1024);
		const literal = stringLiteralLengthFixture(1023);

		expect(physicalLine.source).toHaveLength(1024);
		expect(physicalLine.source).not.toContain('\n');
		expect(literal.metadata.literalLength).toBe(1023);
		expect(literal.source).toContain(`"${'A'.repeat(1023)}"`);
	});

	it('generates exact fixed-string, identifier, and module-name fixtures', () => {
		const fixedString = fixedStringSizeFixture(65527);
		const identifier = identifierLengthFixture(256);
		const moduleName = moduleNameLengthFixture(32);

		expect(fixedString.source).toContain('As String * 65527');
		expect(identifier.metadata.identifierLength).toBe(256);
		expect(String(identifier.metadata.identifier)).toHaveLength(256);
		expect(moduleName.metadata.moduleNameLength).toBe(32);
		expect(String(moduleName.metadata.moduleName)).toHaveLength(32);
	});

	it('generates exact argument-count, array-dimension, and host-call fixtures', () => {
		const argumentCount = argumentCountFixture(61);
		const arrayDimensions = arrayDimensionCountFixture(60);
		const worksheetFunction = excelWorksheetFunctionArgumentCountFixture(256);

		expect(argumentCount.metadata.argumentCount).toBe(61);
		expect(argumentCount.metadata.parameterCount).toBe(61);
		expect(argumentCount.source).toContain('Target 1, 2, 3');
		expect(arrayDimensions.metadata.dimensionCount).toBe(60);
		expect(arrayDimensions.source.match(/\b1\b/g)).toHaveLength(60);
		expect(worksheetFunction.metadata.argumentCount).toBe(256);
		expect(worksheetFunction.source).toContain('Application.WorksheetFunction.Sum(');
	});

	it('rejects invalid generator inputs before producing misleading boundary cases', () => {
		expect(() => continuationCountFixture(-1)).toThrow('count must be non-negative.');
		expect(() => physicalLineLengthFixture(0)).toThrow('length must be positive.');
		expect(() => identifierLengthFixture(1.5)).toThrow('length must be an integer.');
		expect(() => arrayDimensionCountFixture(0)).toThrow('count must be positive.');
	});
});
