// The VBA library's own enums and modules, against its type library.
//
// VBA reads `VbMsgBoxResult.vbYes`, `ColorConstants.vbRed`, `Strings.Left(...)`
// and `Constants.vbCrLf` - measured in Excel 2026-09-18 - and Option Explicit
// called every one of those qualifiers an undeclared variable, where it already
// knew host enums (`XlAxisType.xlCategory`) the same way. The fixture is VBE7.DLL
// read with pythoncom: the model has to match it name for name and value for
// value, so a constant the reference lists but VBA does not have
// (vbUseCompareOption) stays out, and one it has (vbFormMDIForm) stays in.

import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import {
	resolveRuntimeConstant,
	resolveVbaLibraryQualifier,
	VBA_RUNTIME_CONSTANTS,
} from '../src/analyzer/runtime/vbaRuntime';
import { analyzeModule, resolveMemberCompletions } from '../src/analyzer';

interface TypeLibrary {
	enums: Record<string, Record<string, number>>;
	constantModules: Record<string, Record<string, number | string>>;
	functionModules: string[];
}

const library = JSON.parse(readFileSync('tests/fixtures/vbaTypeLibrary.json', 'utf8')) as TypeLibrary;
const lowerNames = (names: Iterable<string>): string[] => [...names].map((name) => name.toLowerCase()).sort();

describe('the VBA runtime constants, against the VBA type library', () => {
	it('has every enum member, with its value and its enum', () => {
		const wrong: string[] = [];
		for (const [enumName, members] of Object.entries(library.enums)) {
			for (const [name, value] of Object.entries(members)) {
				const constant = resolveRuntimeConstant(name);
				if (constant?.value !== value || constant.type !== enumName) {
					wrong.push(`${enumName}.${name}: ${JSON.stringify(constant)}`);
				}
			}
		}
		expect(wrong).toEqual([]);
	});

	it('has every constant of the modules of constants, in its module', () => {
		const wrong: string[] = [];
		for (const [moduleName, members] of Object.entries(library.constantModules)) {
			for (const [name, value] of Object.entries(members)) {
				const constant = resolveRuntimeConstant(name);
				const valueDiffers = typeof value === 'number' && constant?.value !== value;
				if (!constant || constant.module !== moduleName || valueDiffers) {
					wrong.push(`${moduleName}.${name}: ${JSON.stringify(constant)}`);
				}
			}
		}
		expect(wrong).toEqual([]);
	});

	it('puts no constant in a VBA enum the type library does not', () => {
		const claimed = VBA_RUNTIME_CONSTANTS
			.filter((constant) => constant.type && constant.type in library.enums)
			.filter((constant) => !lowerNames(Object.keys(library.enums[constant.type!])).includes(constant.name.toLowerCase()))
			.map((constant) => `${constant.type}.${constant.name}`);
		expect(claimed).toEqual([]);
		// The reference lists it; VBA reads it as an undeclared name.
		expect(resolveRuntimeConstant('vbUseCompareOption')).toBeUndefined();
	});

	it('knows every enum and module as a qualifier, holding exactly its constants', () => {
		for (const [name, members] of [...Object.entries(library.enums), ...Object.entries(library.constantModules)]) {
			const qualifier = resolveVbaLibraryQualifier(name);
			expect(lowerNames((qualifier?.constants ?? []).map((constant) => constant.name)), name)
				.toEqual(lowerNames(Object.keys(members)));
		}
		for (const name of library.functionModules) {
			expect(resolveVbaLibraryQualifier(name), name).toEqual({ name });
		}
	});
});

describe('a VBA enum or module as a qualifier', () => {
	const analyze = (lines: string[]): string[] => analyzeModule(
		['Option Explicit', 'Sub T()', '    Dim v As Variant', ...lines.map((line) => `    ${line}`), '    Debug.Print v', 'End Sub', ''].join('\n'),
		{ host: 'excel', knownIdentifiers: new Set<string>() },
	).map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`);

	it('reads the qualified forms VBA reads', () => {
		expect(analyze([
			'v = VbMsgBoxResult.vbYes',
			'v = VbMsgBoxStyle.vbYesNo',
			'v = FormShowConstants.vbModal',
			'v = VBA.VbMsgBoxResult.vbYes',
			'v = ColorConstants.vbRed',
			'v = KeyCodeConstants.vbKeyReturn',
			'v = SystemColorConstants.vbScrollBars',
			'v = Strings.Left("abc", 1)',
			'v = DateTime.Now',
			'v = Math.Abs(-1)',
		])).toEqual([]);
	});

	it("reads Constants as VBA's module in Excel, which has a Constants enum of its own", () => {
		expect(analyze(['v = Constants.vbCrLf'])).toEqual([]);
	});

	it('reports a member the enum does not have, as the VBE does', () => {
		expect(analyze(['v = VbMsgBoxResult.vbBogus'])).toEqual([
			"member-not-found: Method or data member not found: 'VbMsgBoxResult.vbBogus'.",
		]);
	});

	it('reports vbUseCompareOption, which VBA does not define', () => {
		expect(analyze(['v = vbUseCompareOption'])).toEqual([
			"undeclared-variable: Variable not defined: 'vbUseCompareOption'. Declare it before using it, or remove Option Explicit.",
		]);
	});

	it('offers an enum its own constants after the dot', () => {
		const source = 'Sub T()\n    v = VbMsgBoxResult.\nEnd Sub\n';
		const members = resolveMemberCompletions(source, source.indexOf('VbMsgBoxResult.') + 15, {} as never)
			.map((item) => item.name);
		expect(members.sort()).toEqual(['vbAbort', 'vbCancel', 'vbIgnore', 'vbNo', 'vbOK', 'vbRetry', 'vbYes']);
	});
});
