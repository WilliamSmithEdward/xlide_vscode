// Diagnostics tests: two false positives from the class-member fuzzer of
// issue #414. Measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

function errors(cls: string, ...lines: string[]): string[] {
	const main = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Class1\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
	const classSrc = `Option Explicit\n${cls}\n`;
	const modules = [{ moduleName: 'Module1', source: main }, { moduleName: 'Class1', source: classSrc, type: 'class' }];
	return [
		...analyzeProjectModule(main, modules, 'Module1'),
		...analyzeProjectModule(classSrc, modules, 'Class1', { moduleKind: 'class' }),
	].filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

const get = (type: string): string => `Public Property Get M() As ${type}\nEnd Property`;
const letM = (type: string): string => `Public Property Let M(ByVal v As ${type})\nEnd Property`;

describe('an object Property Get beside a Variant Property Let', () => {
	it('compiles', () => {
		for (const type of ['Collection', 'Object']) {
			expect(errors(`${get(type)}\n${letM('Variant')}`), type).toEqual([]);
		}
	});

	it('leaves the pairs the VBE refuses reported', () => {
		for (const [getType, letType] of [['Long', 'Variant'], ['Variant', 'Long'], ['Long', 'Integer'], ['String', 'Variant'], ['Collection', 'Object']]) {
			expect(errors(`${get(getType)}\n${letM(letType)}`).some((e) => e.startsWith('property-accessor-signature-mismatch')), `${getType}/${letType}`).toBe(true);
		}
	});
});

describe('Set through a Property Set beside a Long Let', () => {
	it('runs', () => {
		const cls = `${get('Long')}\n${letM('Long')}\nPublic Property Set M(ByVal v As Object)\nEnd Property`;
		expect(errors(cls, 'Set c.M = New Collection')).toEqual([]);
	});

	it('is still refused with no Property Set', () => {
		expect(errors(`${get('Long')}\n${letM('Long')}`, 'Set c.M = New Collection').length).toBeGreaterThan(0);
	});
});
