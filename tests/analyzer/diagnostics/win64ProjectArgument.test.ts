// A project's conditional compilation argument `Win64 = 0` (Tools > Project
// Properties) in 64-bit Office turns `#If Win64` off and nothing else.
// Measured in 64-bit Excel, Word, PowerPoint and Access (Microsoft 365 16.0),
// the argument typed into the Project Properties dialog: a Declare without
// PtrSafe is still refused, and VarPtr into a Long and AddressOf into a ByVal
// Long are still a Type mismatch, while the same code with PtrSafe or a
// LongPtr compiles. A module's `#Const Win64 = 0` is the same (issue #1528,
// win64ModuleConst1528.test.ts). So the three rules read the host's Win64,
// and the project's argument decides only the branches.

import { describe, expect, it } from 'vitest';
import { parseProjectConditionalConstants } from '../../../src/analyzer';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';

const HOSTS = ['excel', 'word', 'powerpoint', 'access'] as const;
// Compiled only while #If Win64 is on: "Ambiguous name detected: m".
const PROBE = '#If Win64 Then\nPrivate m As Long\nPrivate m As Long\n#End If\n';

// Each case as measured: its arguments, and the rule that reports it, or
// none where the VBE accepted it.
const CASES: readonly (readonly [name: string, argumentsText: string, rule: string | undefined, body: string])[] = [
	['declare_no_ptrsafe', 'Win64 = 0', 'declare-missing-ptrsafe', 'Private Declare Function GetTickCount Lib "kernel32" () As Long\n'],
	['declare_ptrsafe', 'Win64 = 0', undefined, 'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'],
	['varptr_into_long', 'Win64 = 0', 'longlong-narrowing', 'Sub Main()\n    Dim x As Long\n    Dim n As Long\n    n = VarPtr(x)\nEnd Sub\n'],
	['varptr_into_longptr', 'Win64 = 0', undefined, 'Sub Main()\n    Dim x As Long\n    Dim n As LongPtr\n    n = VarPtr(x)\nEnd Sub\n'],
	['addressof_byval_long', 'Win64 = 0', 'addressof-misuse', 'Sub Take(ByVal v As Long)\nEnd Sub\nSub Callback()\nEnd Sub\nSub Main()\n    Take AddressOf Callback\nEnd Sub\n'],
	['addressof_byval_longptr', 'Win64 = 0', undefined, 'Sub Take(ByVal v As LongPtr)\nEnd Sub\nSub Callback()\nEnd Sub\nSub Main()\n    Take AddressOf Callback\nEnd Sub\n'],
];

// The project as the extension builds it: the arguments text parsed into the
// project's constants, given to the index and to the rules alike.
function errorCodes(body: string, argumentsText: string, host: string): string[] {
	const source = `Option Explicit\n${body}`;
	const project = buildVbaProjectIndex([{ moduleName: 'Module1', source }], undefined, {
		conditionalCompilation: { projectConstants: parseProjectConditionalConstants(argumentsText) },
	});
	return analyzeVbaModuleSource({
		source, moduleName: 'Module1', host,
		...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? '');
}

describe.each(HOSTS)('a project argument Win64 = 0 in 64-bit %s', (host) => {
	it('turns #If Win64 off', () => {
		expect(errorCodes(PROBE, '', host)).toEqual(['duplicate-module-variable']);
		expect(errorCodes(PROBE, 'Win64 = 0', host)).toEqual([]);
	});

	it.each(CASES)('leaves %s as 64-bit Office has it', (_name, argumentsText, rule, body) => {
		expect(errorCodes(body, argumentsText, host)).toEqual(rule ? [rule] : []);
	});
});

describe('a project argument Win64 = 0 in VB6', () => {
	it('leaves the three rules off, as before', () => {
		const rules = new Set(['declare-missing-ptrsafe', 'longlong-narrowing', 'addressof-misuse']);
		for (const [, argumentsText, , body] of CASES) {
			expect(errorCodes(body, argumentsText, 'vb6').filter((code) => rules.has(code))).toEqual([]);
		}
	});
});
