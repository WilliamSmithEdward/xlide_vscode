import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeProjectModule } from './helpers';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import { declaredValueTypeForQualifiedSourceBinding } from '../../src/analyzer/diagnostics/typeInference';

const errors = (source: string) => analyzeModule(source, { knownIdentifiers: new Set<string>() })
	.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code);

const reported: Array<[string, string]> = [
	['Function result', 'Option Explicit\nDefLng C-C\nPublic Function Compute()\nWork Compute\nEnd Function\nPrivate Sub Work(ByRef ComputeResult)\nComputeResult = 7\nEnd Sub\n'],
	['local', 'Option Explicit\nDefLng X-X\nSub Main()\nDim x\nWork x\nDebug.Print x\nEnd Sub\nPrivate Sub Work(ByRef x)\nx = 7\nEnd Sub\n'],
	['array', 'Option Explicit\nDefLng A-Z\nSub Main()\nDim x(1 To 2)\nWork x\nEnd Sub\nPrivate Sub Work(ByRef x())\nx(2) = 7\nEnd Sub\n'],
	['caller parameter', 'Option Explicit\nDefLng A-Z\nSub Main(x)\nWork x\nEnd Sub\nPrivate Sub Work(ByRef x)\nx = 7\nEnd Sub\n'],
];

describe('DefType caller bindings (#1235)', () => {
	for (const newline of ['\n', '\r\n', '\r']) {
		it.each(reported)(`accepts the reported %s with ${JSON.stringify(newline)} line endings`, (_name, source) => {
			expect(errors(source.replaceAll('\n', newline))).toEqual([]);
		});
	}

	it.each(['Byte', 'Int', 'Lng', 'Sng', 'Dbl', 'Cur', 'Date', 'Str', 'Obj'])(
		'uses Def%s for scalar locals, array elements and module variables', (kind) => {
			const source = `Option Explicit\nDef${kind} X-X\nDim xModule\nSub Main()\nDim x\nDim xArray(1 To 2)\nWork x\nWork xArray(1)\nWork xModule\nEnd Sub\nPrivate Sub Work(ByRef x)\nEnd Sub\n`;
			expect(errors(source)).toEqual([]);
		},
	);

	it('uses the Property Get result type', () => {
		expect(errors('Option Explicit\nDefLng C-C\nPublic Property Get Compute()\nWork Compute\nEnd Property\nPrivate Sub Work(ByRef ComputeResult)\nComputeResult = 7\nEnd Sub\n')).toEqual([]);
	});

	it.each(['Dim x As Variant', 'Dim x As String', 'Dim x$'])(
		'keeps an explicit local type: %s', (declaration) => {
			expect(errors(`Option Explicit\nDefLng X-X\nSub Main()\n${declaration}\nWork x\nEnd Sub\nPrivate Sub Work(ByRef x)\nEnd Sub\n`)).toContain('byref-argument-type-mismatch');
		},
	);

	it('keeps an explicit array element type', () => {
		expect(errors('Option Explicit\nDefLng A-Z\nSub Main()\nDim x(1 To 2) As Variant\nWork x\nEnd Sub\nPrivate Sub Work(ByRef x())\nEnd Sub\n')).toContain('argument-shape-mismatch');
	});

	it('keeps a ParamArray as Variant elements despite DefLng', () => {
		expect(errors('Option Explicit\nDefLng A-Z\nSub Main(ParamArray x())\nWork x\nEnd Sub\nPrivate Sub Work(ByRef x() As Long)\nEnd Sub\n')).toContain('argument-shape-mismatch');
	});

	it('keeps names outside the DefType range as Variant', () => {
		expect(errors('Option Explicit\nDefLng X-X\nSub Main()\nDim y\nWork y\nEnd Sub\nPrivate Sub Work(ByRef x)\nEnd Sub\n')).toContain('byref-argument-type-mismatch');
	});

	it('resolves a qualified variable using its owning module default type', () => {
		const source = 'Option Explicit\nDefLng X-X\nDim x\nSub Main()\nWork Caller.x\nEnd Sub\nPrivate Sub Work(ByRef x)\nEnd Sub\n';
		expect(analyzeModule(source, { moduleName: 'Caller', knownIdentifiers: new Set(['caller']) })
			.filter((diagnostic) => diagnostic.severity === 'error')).toEqual([]);
	});

	it('does not apply caller DefLng to a foreign Variant', () => {
		const argument = 'x';
		const source = `Option Explicit\nDefLng X-X\nSub Main()\nWork ${argument}\nEnd Sub\nPrivate Sub Work(ByRef x)\nEnd Sub\n`;
		const diagnostics = analyzeProjectModule(source, [
			{ moduleName: 'Caller', source },
			{ moduleName: 'Other', source: 'Option Explicit\nPublic x\n' },
		], 'Caller', { knownIdentifiers: new Set<string>() });
		expect(diagnostics.filter((diagnostic) => diagnostic.code === 'byref-argument-type-mismatch')).toHaveLength(1);
	});

	it('uses the implicit local type when it shadows an explicitly typed module variable', () => {
		expect(errors('Option Explicit\nDefLng X-X\nDim x As String\nSub Main()\nDim x\nWork x\nEnd Sub\nPrivate Sub Work(ByRef x)\nEnd Sub\n')).toEqual([]);
	});

	it('keeps an explicit Function result type', () => {
		expect(errors('Option Explicit\nDefLng C-C\nPublic Function Compute() As Variant\nWork Compute\nEnd Function\nPrivate Sub Work(ByRef ComputeResult)\nEnd Sub\n')).toContain('byref-argument-type-mismatch');
	});
	it('does not apply caller DefLng when resolving a qualified foreign declaration', () => {
		const caller = buildModuleSymbols('Caller', 'standard', 'DefLng X-X\n');
		const other = buildModuleSymbols('Other', 'standard', 'Public x\n');
		expect(declaredValueTypeForQualifiedSourceBinding(caller, other.root.children, 'Other', 'x'))
			.toEqual({ resolved: true, asType: undefined });
	});
});
