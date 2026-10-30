// The analyzer functions that build their own symbols, for editor features and
// refactors, take the project's conditional compilation: under 32-bit Office
// (Win64 off) the #Else arm of #If Win64 is the code, and its declarations
// are the ones a hover, a completion or a refactor sees.

import { describe, expect, it } from 'vitest';
import {
	analyzeModule,
	introduceParameter,
	moveToModule,
	resolveExpressionType,
	resolveSignatureHelp,
	resolveHover,
	type ConditionalCompilationEnvironment,
} from '../../src/analyzer';
import { resolveAssignmentValueCompletion, resolveEnumValues } from '../../src/analyzer/completion/assignmentValueCompletion';
import { createIdentifierCompletionResolver } from '../../src/analyzer/completion/identifierCompletion';
import { conditionalCompilationKey } from '../../src/analyzer/conditional/conditionalCompilation';
import { editorModuleSymbols } from '../../src/analyzer/symbols/editorModuleSymbols';
import { checkMissingLibraryReference, librariesNamedIn } from '../../src/analyzer/diagnostics/rules/missingReference';
import { getExcelObjectModel } from '../../src/analyzer/host/excelObjectModel';
import { referencedProjectSurface } from '../../src/analyzer/symbols/referencedProjects';

const OFFICE_32: ConditionalCompilationEnvironment = { projectConstants: {}, compilerConstants: { Win64: 0 } };
const SOURCE = [
	'Option Explicit',
	'#If Win64 Then',
	'Public Wide As LongLong',
	'#Else',
	'Public Narrow As Long',
	'Public Enum Width32',
	'    W32 = 1',
	'End Enum',
	'#End If',
	'Sub Main()',
	'    Dim v As Variant',
	'    v = Narrow',
	'End Sub',
	'',
].join('\n');

const names = (env?: ConditionalCompilationEnvironment) =>
	(editorModuleSymbols('Module1', 'standard', SOURCE, env).root.children ?? []).map((symbol) => symbol.name);

describe('the cache key of a conditional compilation', () => {
	it('ignores the order of the constants, and tells absent project arguments from none', () => {
		expect(conditionalCompilationKey({ compilerConstants: { Win64: 0, VBA7: 1 } }))
			.toBe(conditionalCompilationKey({ compilerConstants: { VBA7: 1, Win64: 0 } }));
		expect(conditionalCompilationKey(undefined)).toBe(conditionalCompilationKey({}));
		expect(conditionalCompilationKey({ projectConstants: {} })).not.toBe(conditionalCompilationKey({}));
		expect(conditionalCompilationKey(OFFICE_32)).not.toBe(conditionalCompilationKey({ projectConstants: {} }));
	});

	it('keys names as they are read: case-blind, the last value winning', () => {
		const first = { projectConstants: { Flag: 0, flag: 1 } };
		const second = { projectConstants: { flag: 1, Flag: 0 } };
		expect(conditionalCompilationKey(first)).not.toBe(conditionalCompilationKey(second));
		expect(conditionalCompilationKey({ projectConstants: { FLAG: 1 } })).toBe(conditionalCompilationKey(first));
		const source = 'Option Explicit\n#If Flag Then\nPublic OnArm As Long\n#Else\nPublic OffArm As Long\n#End If\n';
		const declared = (env: ConditionalCompilationEnvironment) =>
			(editorModuleSymbols('Flags', 'standard', source, env).root.children ?? []).map((symbol) => symbol.name);
		expect(declared(first)).toEqual(['OnArm']);
		expect(declared(second)).toEqual(['OffArm']);
	});

	it('keys every value apart, Infinity, -Infinity and NaN too', () => {
		const keys = [Infinity, -Infinity, NaN, 0, 1, '1', true].map((Flag) => conditionalCompilationKey({ projectConstants: { Flag } }));
		expect(new Set(keys).size).toBe(keys.length);
		const source = 'Option Explicit\n#If Flag > 0 Then\nPublic Above As Long\n#Else\nPublic Below As Long\n#End If\n';
		const declared = (Flag: number) =>
			(editorModuleSymbols('Bounds', 'standard', source, { projectConstants: { Flag } }).root.children ?? []).map((symbol) => symbol.name);
		expect(declared(Infinity)).toEqual(['Above']);
		expect(declared(-Infinity)).toEqual(['Below']);
	});

	it('keys compiler constants as they are read over the defaults', () => {
		// The default Win64 keeps its place in the merge, so win64 = 1 is read last here.
		const merged = { compilerConstants: { win64: 1, Win64: 0 } };
		expect(conditionalCompilationKey(merged)).not.toBe(conditionalCompilationKey(OFFICE_32));
		expect(conditionalCompilationKey(merged)).toBe(conditionalCompilationKey({}));
		expect(conditionalCompilationKey({ compilerConstants: { Win64: 1 } })).toBe(conditionalCompilationKey(undefined));
		const source = 'Option Explicit\n#If Win64 Then\nPublic Wide As LongLong\n#Else\nPublic Narrow As Long\n#End If\n';
		const declared = (env: ConditionalCompilationEnvironment) =>
			(editorModuleSymbols('Merged', 'standard', source, env).root.children ?? []).map((symbol) => symbol.name);
		expect(declared(OFFICE_32)).toEqual(['Narrow']);
		expect(declared(merged)).toEqual(['Wide']);
	});
});

describe('editor symbols under a conditional compilation', () => {
	it('declare what the arm compiled for that Office declares, cached apart', () => {
		expect(names()).toContain('Wide');
		expect(names()).not.toContain('Narrow');
		expect(names(OFFICE_32)).toEqual(expect.arrayContaining(['Narrow', 'Width32']));
		expect(names(OFFICE_32)).not.toContain('Wide');
		// Same source, same module: the cache holds both answers.
		expect(names()).toContain('Wide');
		expect(names({})).toContain('Wide');
	});

	it('reach identifier completion', () => {
		const offset = SOURCE.indexOf('Narrow\nEnd Sub') + 'Nar'.length;
		const labels = (env?: ConditionalCompilationEnvironment) =>
			createIdentifierCompletionResolver(SOURCE, env ? { conditionalCompilation: env } : {})(offset).map((item) => item.name);
		expect(labels(OFFICE_32)).toContain('Narrow');
		expect(labels()).not.toContain('Narrow');
	});

	it('reach hover', () => {
		const offset = SOURCE.indexOf('Narrow\nEnd Sub') + 1;
		expect(resolveHover(SOURCE, offset, { conditionalCompilation: OFFICE_32 })?.signature).toMatch(/Narrow As Long/);
		expect(resolveHover(SOURCE, offset)?.signature ?? '').not.toMatch(/Narrow As Long/);
	});

	it('reach expression types', () => {
		const start = SOURCE.indexOf('Narrow\nEnd Sub');
		const span = { start, end: start + 'Narrow'.length };
		expect(resolveExpressionType(SOURCE, span, { conditionalCompilation: OFFICE_32 })?.type).toBe('Long');
		expect(resolveExpressionType(SOURCE, span)?.type).not.toBe('Long');
	});

	it('reach the enum values a source assignment offers', () => {
		expect(resolveEnumValues(SOURCE, 'Width32', { conditionalCompilation: OFFICE_32 })?.enumName).toBe('Width32');
		expect(resolveEnumValues(SOURCE, 'Width32', {})).toBeUndefined();
	});

	it('reach the assignment target an assignment completion types', () => {
		// Win64 on, Mode is a String; off, a VbMsgBoxStyle, whose values are offered.
		const source = 'Option Explicit\n#If Win64 Then\nPublic Mode As String\n#Else\nPublic Mode As VbMsgBoxStyle\n#End If\nSub M()\n    Mode = \nEnd Sub\n';
		const offset = source.indexOf('Mode = ') + 'Mode = '.length;
		expect(resolveAssignmentValueCompletion(source, offset, { conditionalCompilation: OFFICE_32 })?.enumName).toBe('VbMsgBoxStyle');
		expect(resolveAssignmentValueCompletion(source, offset)).toBeUndefined();
	});

	it('reach a referenced project surface', () => {
		const exported = (env?: ConditionalCompilationEnvironment) => referencedProjectSurface('Library', [{ name: 'Module1', kind: 'standard', source: SOURCE }], env)
			.modules[0].members;
		expect(exported(OFFICE_32)).toContain('Narrow');
		expect(exported(OFFICE_32)).not.toContain('Wide');
		expect(exported()).toContain('Wide');
	});
});

describe('signature help under a conditional compilation', () => {
	it('shows the declaration the arm compiled for that Office declares', () => {
		const source = 'Option Explicit\n#If Win64 Then\nSub Work(ByVal wide As LongLong)\nEnd Sub\n#Else\nSub Work(ByVal narrow As Long)\nEnd Sub\n#End If\nSub Main()\n    Work 1\nEnd Sub\n';
		const offset = source.indexOf('Work 1') + 'Work '.length;
		const label = (env: ConditionalCompilationEnvironment) => resolveSignatureHelp(source, offset, { conditionalCompilation: env })?.label;
		expect(label(OFFICE_32)).toMatch(/narrow As Long/);
		expect(label({ projectConstants: {} })).toMatch(/wide As LongLong/);
		// With no environment, as every caller passes today: 64-bit Office.
		expect(resolveSignatureHelp(source, offset)?.label).toMatch(/wide As LongLong/);
		// Same source: each answer is cached apart, and a project argument decides #If.
		expect(label(OFFICE_32)).toMatch(/narrow As Long/);
		expect(label({ projectConstants: { Win64: 0 } })).toMatch(/narrow As Long/);
	});
});

describe('signature help with no conditional compilation', () => {
	it('looks only at an arm the default Office compiles, as hover does', () => {
		// #If Mac is off in Windows Office, so the #Else declaration is the one compiled.
		const source = 'Option Explicit\n#If Mac Then\nSub Work(ByVal m As String)\nEnd Sub\n#Else\nSub Work(ByVal w As Long)\nEnd Sub\n#End If\nSub Main()\n    Work 1\nEnd Sub\n';
		expect(resolveSignatureHelp(source, source.indexOf('Work 1') + 'Work '.length)?.label).toMatch(/w As Long/);
	});
});

describe('library qualifiers under a conditional compilation', () => {
	it('see a module variable only where its arm is compiled', () => {
		// Win64 on, Excel is the module's own variable; off, it names the library.
		const source = 'Option Explicit\n#If Win64 Then\nPrivate Excel As Object\n#End If\nSub Main()\n    Excel.Calculate\nEnd Sub\n';
		expect([...librariesNamedIn(source, undefined, new Set(), OFFICE_32)]).toEqual(['excel']);
		expect([...librariesNamedIn(source)]).toEqual([]);
	});

	it('report a missing library only where its arm is compiled', () => {
		// Win64 on, Word is the module's own variable; off, it names a library Excel does not reference.
		const source = 'Option Explicit\n#If Win64 Then\nPrivate Word As Object\n#End If\nSub Main()\n    Word.Activate\nEnd Sub\n';
		const reported = (env?: ConditionalCompilationEnvironment) => {
			const codes: string[] = [];
			checkMissingLibraryReference(source, getExcelObjectModel(), (code) => codes.push(code), new Set(), undefined, undefined, env);
			return codes;
		};
		expect(reported(OFFICE_32)).toEqual(['missingLibraryReference']);
		expect(reported()).toEqual([]);
	});
});

describe('refactors under a conditional compilation', () => {
	it('bind a caller in another module as that Office compiles it: Introduce Parameter', () => {
		const source = 'Option Explicit\nPublic Sub Work()\n    Dim limit As Long\n    limit = 5\n    Debug.Print limit\nEnd Sub\n';
		// Win64 on, the caller's Work is its own; off, it is M.Work, which gains the argument.
		const other = 'Option Explicit\n#If Win64 Then\nPrivate Sub Work()\nEnd Sub\n#End If\nSub Caller()\n    Work\nEnd Sub\n';
		const callerEdits = (env?: ConditionalCompilationEnvironment) => {
			const result = introduceParameter({ source, offset: source.indexOf('limit As'), moduleName: 'M', otherModuleSources: { Other: other }, conditionalCompilation: env });
			return result.ok ? (result.otherModules ?? []).filter((mod) => mod.moduleName === 'Other').flatMap((mod) => mod.edits.map((edit) => edit.newText)) : ['refused'];
		};
		expect(callerEdits()).toEqual([]);
		expect(callerEdits(OFFICE_32)).toEqual([' 5']);
	});

	it('bind a caller in another module as that Office compiles it: Move to Module', () => {
		const source = 'Option Explicit\nPublic Sub Work()\nEnd Sub\n';
		// Win64 on, M is the caller's own field; off, M names the module, which moves.
		const caller = 'Option Explicit\n#If Win64 Then\nPrivate M As Object\n#End If\nSub Caller()\n    M.Work\nEnd Sub\n';
		const callerEdits = (env?: ConditionalCompilationEnvironment) => {
			const result = moveToModule({ source, offset: source.indexOf('Work'), moduleName: 'M', targetModuleName: 'N',
				otherModuleSources: { N: 'Option Explicit\n', C: caller }, conditionalCompilation: env });
			return result.ok ? (result.otherModules ?? []).filter((mod) => mod.moduleName === 'C').flatMap((mod) => mod.edits.map((edit) => edit.newText)) : ['refused'];
		};
		expect(callerEdits()).toEqual([]);
		expect(callerEdits(OFFICE_32)).toEqual(['N']);
	});
});

describe('the declared type an undeclared variable is offered', () => {
	it('comes from the arm compiled for that Office', () => {
		const source = 'Option Explicit\n#If Win64 Then\nPublic Wide As LongLong\n#Else\nPublic Narrow As Long\n#End If\nSub Main()\n    fresh = Narrow\nEnd Sub\n';
		const offered = (env: ConditionalCompilationEnvironment) => analyzeModule(source, { conditionalCompilation: env, knownIdentifiers: new Set<string>() })
			.filter((diagnostic) => diagnostic.code === 'undeclared-variable' && diagnostic.message.includes("'fresh'"))
			.map((diagnostic) => (diagnostic.data as { declareVariable?: { declaredType?: string } } | undefined)?.declareVariable?.declaredType);
		expect(offered(OFFICE_32)).toEqual(['Long']);
		expect(offered({ projectConstants: {} })).toEqual(['Variant']);
	});
});
