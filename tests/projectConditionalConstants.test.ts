// The VBE's "Conditional Compilation Arguments" project property.
//
// XLIDE knew the compiler constants (VBA7, Win64, Mac) but never the project's
// own, so `#If MY_FLAG Then` was undecidable in every project - including ones
// that declare MY_FLAG in the property sheet and would compile only one arm.
// MS-OVBA keeps the property in the dir stream as record 0x000C, with a Unicode
// twin at 0x003C.

import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { parseProjectConditionalConstants } from '../src/analyzer';
import { readModules } from '../src/vba/projectService';
import { analyzeModule } from '../src/analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../src/vbaProjectAnalysis';

const CORPUS = path.resolve('..', 'xlide_vscode_testing');

describe('parsing the project property', () => {
	it('reads `Name = Value` pairs separated by colons', () => {
		expect(parseProjectConditionalConstants('stdFullIntegration = 1')).toEqual({
			stdFullIntegration: 1,
		});
		expect(parseProjectConditionalConstants('Debugging=1:Trace = 2')).toEqual({
			Debugging: 1,
			Trace: 2,
		});
	});

	it('keeps the VBA boolean encoding, where True is -1', () => {
		expect(parseProjectConditionalConstants('A = -1 : B = 0')).toEqual({ A: -1, B: 0 });
	});

	it('skips an entry that names nothing rather than guessing', () => {
		expect(parseProjectConditionalConstants('')).toEqual({});
		expect(parseProjectConditionalConstants(undefined)).toEqual({});
		expect(parseProjectConditionalConstants('Bad')).toEqual({});
		expect(parseProjectConditionalConstants('9Bad = 1')).toEqual({});
	});

	it('keeps a non-integer value as its text instead of coercing it', () => {
		expect(parseProjectConditionalConstants('Label = dev')).toEqual({ Label: 'dev' });
	});
});

describe('reading the property out of a real project', () => {
	const corpusPresent = fs.existsSync(path.join(CORPUS, 'fullBuild.xlsm'));

	it.runIf(corpusPresent)('finds the arguments a workbook declares', () => {
		const entries = readModules(path.join(CORPUS, 'fullBuild.xlsm'));
		const raw = entries.find((e) => e.projectConditionalConstants)?.projectConditionalConstants;
		expect(raw).toBe('stdFullIntegration = 1');
		expect(parseProjectConditionalConstants(raw)).toEqual({ stdFullIntegration: 1 });
	});

	it.runIf(corpusPresent)('reports none for a project that declares none', () => {
		const entries = readModules(path.join(CORPUS, 'fastjson.xlsm'));
		expect(entries.find((e) => e.projectConditionalConstants)).toBeUndefined();
	});
});

describe('supplying the arguments decides a branch', () => {
	const SRC = [
		'Option Explicit',
		'#If stdFullIntegration Then',
		'Dim Wanted As Long',
		'#Else',
		'Dim Unwanted As Long',
		'#End If',
		'Sub T()',
		'    Wanted = 1',
		'End Sub',
	].join('\r\n') + '\r\n';

	/**
	 * Analyzes through the real project path, so the symbol table and the rules
	 * are built under the SAME constants. Supplying them only to the rules would
	 * leave the losing arm's declarations in the symbol table and hide the
	 * effect entirely.
	 */
	const undeclared = (source: string, raw?: string): number => {
		const project = buildVbaProjectIndex(
			[{ moduleName: 'M', source }],
			undefined,
			raw
				? { conditionalCompilation: { projectConstants: parseProjectConditionalConstants(raw) } }
				: {},
		);
		const options = projectAnalysisOptionsForModule(project, 'M', projectProcedureSignatures(project));
		return analyzeModule(source, { moduleName: 'M', ...options })
			.filter((d) => d.code === 'undeclared-variable').length;
	};

	it('leaves both arms live when the constant is unknown', () => {
		// Both declarations are visible, so neither use is undeclared.
		expect(undeclared(SRC)).toBe(0);
		expect(undeclared(SRC.replace('Wanted = 1', 'Unwanted = 1'))).toBe(0);
	});

	it('drops the losing arm once the project declares the constant', () => {
		expect(undeclared(SRC, 'stdFullIntegration = 1')).toBe(0);
		// `Unwanted` sits in the arm that loses, so it is no longer declared at
		// all and using it is the error the VBE would raise.
		expect(undeclared(SRC.replace('Wanted = 1', 'Unwanted = 1'), 'stdFullIntegration = 1')).toBe(1);
	});
});
