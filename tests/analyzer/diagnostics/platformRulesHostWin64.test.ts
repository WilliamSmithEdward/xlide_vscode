// Diagnostics tests: the platform rules follow the Office the code runs in,
// the host's Win64, and never a project's own Win64. Measured through the VBE
// oracle in Excel, Word, PowerPoint and Access, 64-bit Microsoft 365 16.0
// (build 20527) and 32-bit Office 2016 (build 16.0.12527.22286), the four
// hosts agreeing on every case:
//
// - A project's conditional compilation argument `Win64 = 0` (Tools >
//   Project Properties), or a module's `#Const Win64 = 0`, in 64-bit Office
//   turns `#If Win64` off and nothing else: a Declare without PtrSafe is
//   still refused, and VarPtr into a Long and AddressOf into a ByVal Long are
//   still Type mismatch.
// - `Win64 = 1`, as an argument or a `#Const`, in 32-bit Office turns
//   `#If Win64` on and nothing else: those three still compile.

import { describe, expect, it } from 'vitest';
import { analyzeModule, type ConditionalCompilationEnvironment } from '../../../src/analyzer';

const SHAPES = {
	declare: ['declare-missing-ptrsafe', 'Private Declare Function GetTickCount Lib "kernel32" () As Long\nSub Main()\n    Dim n As Long\n    n = GetTickCount()\nEnd Sub\n'],
	varptr: ['longlong-narrowing', 'Sub Main()\n    Dim x As Long\n    Dim n As Long\n    n = VarPtr(x)\nEnd Sub\n'],
	addressOf: ['addressof-misuse', 'Sub Take(ByVal v As Long)\nEnd Sub\nSub Callback()\nEnd Sub\nSub Main()\n    Take AddressOf Callback\nEnd Sub\n'],
} as const;
// The #If Win64 arm, when it is compiled, declares a local twice.
const ARM = 'Sub Main()\n#If Win64 Then\n    Dim a As Long\n    Dim a As Long\n#End If\nEnd Sub\n';
const HOSTS = ['excel', 'word', 'powerpoint', 'access'] as const;
const OFFICE_64: ConditionalCompilationEnvironment = { projectConstants: {} };
const OFFICE_32: ConditionalCompilationEnvironment = { projectConstants: {}, compilerConstants: { Win64: 0 } };

function codes(body: string, host: string, conditionalCompilation: ConditionalCompilationEnvironment, moduleConst = ''): string[] {
	const source = `${moduleConst}Option Explicit\n${body}`;
	return analyzeModule(source, { conditionalCompilation, ...(host === 'excel' ? {} : { host }) })
		.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? '');
}

function withProject(env: ConditionalCompilationEnvironment, win64: number): ConditionalCompilationEnvironment {
	return { ...env, projectConstants: { Win64: win64 } };
}

describe.each(HOSTS)('the platform rules in %s', (host) => {
	it.each(Object.entries(SHAPES))('report %s as 64-bit Office, whatever the project or the module says', (_name, [code, body]) => {
		expect(codes(body, host, OFFICE_64)).toEqual([code]);
		expect(codes(body, host, withProject(OFFICE_64, 0))).toEqual([code]);
		expect(codes(body, host, OFFICE_64, '#Const Win64 = 0\n')).toEqual([code]);
	});

	it.each(Object.entries(SHAPES))('leave %s alone as 32-bit Office, whatever the project or the module says', (_name, [, body]) => {
		expect(codes(body, host, OFFICE_32)).toEqual([]);
		expect(codes(body, host, withProject(OFFICE_32, 1))).toEqual([]);
		expect(codes(body, host, OFFICE_32, '#Const Win64 = 1\n')).toEqual([]);
	});

	it('let the project and the module decide #If Win64', () => {
		const armed = (env: ConditionalCompilationEnvironment, moduleConst = '') =>
			codes(ARM, host, env, moduleConst).includes('duplicate-declaration');
		expect(armed(OFFICE_64)).toBe(true);
		expect(armed(withProject(OFFICE_64, 0))).toBe(false);
		expect(armed(OFFICE_64, '#Const Win64 = 0\n')).toBe(false);
		expect(armed(OFFICE_32)).toBe(false);
		expect(armed(withProject(OFFICE_32, 1))).toBe(true);
		expect(armed(OFFICE_32, '#Const Win64 = 1\n')).toBe(true);
	});
});

describe('the platform rules in VB6', () => {
	it('stand aside, as before', () => {
		expect(analyzeModule(`Option Explicit\n${SHAPES.declare[1]}`, { host: 'vb6' })
			.filter((diagnostic) => diagnostic.code === 'declare-missing-ptrsafe')).toEqual([]);
		expect(analyzeModule(`Option Explicit\n${SHAPES.varptr[1]}`, { host: 'vb6' })
			.filter((diagnostic) => diagnostic.code === 'longlong-narrowing')).toEqual([]);
	});
});
