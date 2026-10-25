// Casing of conditional-compilation names: what the VBE prints for `#If`,
// `#ElseIf` and `#Const` lines, and for a code line that names a `#Const`.
// Every expectation is what Excel, Word, PowerPoint and Access 16.0 showed
// for a file holding the input, read back from CodeModule.Lines; the
// formatter only adds its indentation.

import { describe, expect, it } from 'vitest';
import {
	conditionalConstantDeclarations,
	conditionalDeclarationScanCountForTests,
	conditionalNameCase,
} from '../../src/analyzer/format/conditionalNameCase';
import { formatVbaModule } from '../../src/analyzer/format/formatModule';

interface Project {
	/** Every module's lines, in project order. */
	modules: string[][];
	/** The project's Conditional Compilation Arguments. */
	arguments?: string;
	/** Names an ordinary declaration spells, which the provider answers before these. */
	declared?: string[];
}

/** Formats `project.modules[index]` with the provider's order: ordinary declarations, then these names. */
function format(project: Project, index = 0): string[] {
	const sources = project.modules.map((lines) => lines.join('\n'));
	const conditional = conditionalNameCase(sources.map(conditionalConstantDeclarations), project.arguments);
	const declared = new Map((project.declared ?? []).map((name) => [name.toLowerCase(), name]));
	const result = formatVbaModule(sources[index], {
		tabSize: 4,
		insertSpaces: true,
		identifierCase: (name) => declared.get(name.toLowerCase()) ?? conditional(name),
	});
	expect(result.refusal).toBeUndefined();
	return (result.text as string).split('\n');
}

const one = (...lines: string[]): Project => ({ modules: [lines] });

describe('formatVbaModule - conditional-compilation names', () => {
	it("spells VBA's compiler constants the VBE's way, whatever they were written as", () => {
		expect(format(one(
			'#If vba7 Then',
			'#ElseIf VBA6 Then',
			'#ElseIf WIN16 Then',
			'#ElseIf win32 Then',
			'#ElseIf wIN64 Or Win32 Then',
			'#ElseIf vBa7 And Not mAc Then',
			'#End If',
		))).toEqual([
			'#If VBA7 Then',
			'#ElseIf VBA6 Then',
			'#ElseIf Win16 Then',
			'#ElseIf Win32 Then',
			'#ElseIf Win64 Or Win32 Then',
			'#ElseIf VBA7 And Not Mac Then',
			'#End If',
		]);
	});

	it('spells them so on a code line too, and on directive lines inside a procedure', () => {
		expect(format(one('Sub T()', 'x = win64', '#If vba7 Then', 'y = MAC', '#End If', 'End Sub'))).toEqual(
			['Sub T()', '    x = Win64', '    #If VBA7 Then', '        y = Mac', '    #End If', 'End Sub'],
		);
	});

	it('leaves TWINBASIC as written: VBA does not define it', () => {
		expect(format(one('#If twinbasic Then', '#End If'))).toEqual(['#If twinbasic Then', '#End If']);
	});

	it('gives a #Const name the spelling of its #Const, before and after it', () => {
		expect(format(one(
			'#If LateFlag Then',
			'#End If',
			'#Const DebugMode = 1',
			'#If debugmode Or DEBUGMODE = 1 Then',
			'#End If',
			'#If LATEFLAG Then',
			'#End If',
			'#Const lateflag = 1',
		))).toEqual([
			'#If lateflag Then',
			'#End If',
			'#Const DebugMode = 1',
			'#If DebugMode Or DebugMode = 1 Then',
			'#End If',
			'#If lateflag Then',
			'#End If',
			'#Const lateflag = 1',
		]);
	});

	it('takes the last #Const of a name, and spells the earlier #Const lines that way too', () => {
		expect(format(one(
			'#If VBA7 Then',
			'#Const InBranch = 1',
			'#Else',
			'#Const INBRANCH = 0',
			'#End If',
			'#If inbranch Then',
			'#End If',
		))).toEqual([
			'#If VBA7 Then',
			'    #Const INBRANCH = 1',
			'#Else',
			'    #Const INBRANCH = 0',
			'#End If',
			'#If INBRANCH Then',
			'#End If',
		]);
	});

	it('counts a #Const in a branch that is not compiled, and one inside a procedure', () => {
		expect(format(one(
			'#If False Then',
			'#Const SkipOnly = 1',
			'#End If',
			'#If skiponly Then',
			'#End If',
			'Sub T()',
			'#Const InProc = 1',
			'#If inproc Then',
			'#End If',
			'End Sub',
		))).toEqual([
			'#If False Then',
			'    #Const SkipOnly = 1',
			'#End If',
			'#If SkipOnly Then',
			'#End If',
			'Sub T()',
			'    #Const InProc = 1',
			'    #If InProc Then',
			'    #End If',
			'End Sub',
		]);
	});

	it("spells a name another module's #Const declares, the last module in project order winning", () => {
		const project: Project = {
			modules: [
				['#Const SharedFlag = 1', '#If sharedflag Then', '#End If'],
				['#Const SHAREDFLAG = 1', '#If SharedFlag Then', '#End If'],
				['#If onlyhere Or sharedflag Then', '#End If'],
				['#Const OnlyHere = 1'],
			],
		};
		expect(format(project, 0)).toEqual(['#Const SHAREDFLAG = 1', '#If SHAREDFLAG Then', '#End If']);
		expect(format(project, 1)).toEqual(['#Const SHAREDFLAG = 1', '#If SHAREDFLAG Then', '#End If']);
		expect(format(project, 2)).toEqual(['#If OnlyHere Or SHAREDFLAG Then', '#End If']);
	});

	it("spells a name the project's arguments declare, a #Const of it winning", () => {
		const project: Project = {
			modules: [['#If projarg Or PCASE Or argone Or ARGTWO Then', '#End If', '#Const PCASE = 2']],
			arguments: 'ProjArg = 1 : PCase = 1 : ArgOne = 1 : argTwo = 0',
		};
		expect(format(project)).toEqual(['#If ProjArg Or PCASE Or ArgOne Or argTwo Then', '#End If', '#Const PCASE = 2']);
	});

	it('spells a code line that names a #Const the way the #Const does', () => {
		expect(format(one('#Const CodeUse = 1', 'Sub T()', 'Debug.Print codeuse', 'End Sub'))).toEqual(
			['#Const CodeUse = 1', 'Sub T()', '    Debug.Print CodeUse', 'End Sub'],
		);
	});

	it('lets a #Const or an ordinary declaration of a compiler constant respell it', () => {
		expect(format(one('#Const WIN64 = 0', '#If win64 Or Win64 Then', '#End If'))).toEqual(
			['#Const WIN64 = 0', '#If WIN64 Or WIN64 Then', '#End If'],
		);
		// `Public mac As Long` and `Public WIN64 As Long` in another module.
		expect(format({ modules: [['#If Mac Then', '#ElseIf win64 Then', '#End If']], declared: ['mac', 'WIN64'] })).toEqual(
			['#If mac Then', '#ElseIf WIN64 Then', '#End If'],
		);
	});

	it('lets an ordinary declaration spell a name the arguments or a #Const also declare, when it comes later', () => {
		// `Dim VARARG As Long` beside `VarArg = 1` in the arguments, and
		// `Dim MyFlag4 As Long` after `#Const myflag4 = 1`.
		expect(format({
			modules: [['Dim VARARG As Long', '#If vararg Then', '#End If', '#Const myflag4 = 1', 'Dim MyFlag4 As Long', '#If MYFLAG4 Then', '#End If']],
			arguments: 'VarArg = 1',
			declared: ['VARARG', 'MyFlag4'],
		})).toEqual(['Dim VARARG As Long', '#If VARARG Then', '#End If', '#Const MyFlag4 = 1', 'Dim MyFlag4 As Long', '#If MyFlag4 Then', '#End If']);
	});

	it('cases and spaces the rest of a directive line the way the VBE does', () => {
		expect(format(one(
			'#const opa=1',
			'#const opb = "abc"',
			'#const opc = true',
			'#const ope=opa+2',
			'#if opa=1 and opb="abc" then',
			'#elseif not opc or opa<>2 xor opa>=1 then',
			'#elseif opa+1=2 eqv false then',
			'#elseif opa mod 2 = 1 imp true then',
			'#elseif opa-1 then',
			'#elseif -opa then',
			'#else',
			'#end if',
			'#If (opa) Then',
			'#End If',
		))).toEqual([
			'#Const opa = 1',
			'#Const opb = "abc"',
			'#Const opc = True',
			'#Const ope = opa + 2',
			'#If opa = 1 And opb = "abc" Then',
			'#ElseIf Not opc Or opa <> 2 Xor opa >= 1 Then',
			'#ElseIf opa + 1 = 2 Eqv False Then',
			'#ElseIf opa Mod 2 = 1 Imp True Then',
			'#ElseIf opa - 1 Then',
			'#ElseIf -opa Then',
			'#Else',
			'#End If',
			'#If (opa) Then',
			'#End If',
		]);
	});

	it('counts a #Const continued after the #, or with Rem or a colon after it, but not one after a statement', () => {
		expect(format(one(
			'# _',
			'Const ContName = 1',
			'#Const RemNote = 1 Rem note',
			'#Const ColonEnd = 1:',
			'Sub S()',
			'x = 1: #Const AfterColon = 1',
			'End Sub',
			'#If contname Or remnote Or colonend Or aftercolon Then',
			'#End If',
		))).toEqual([
			'# _',
			'    Const ContName = 1',
			'#Const RemNote = 1 Rem note',
			'#Const ColonEnd = 1:',
			'Sub S()',
			'    x = 1: #Const AfterColon = 1',
			'End Sub',
			'#If ContName Or RemNote Or ColonEnd Or aftercolon Then',
			'#End If',
		]);
	});

	it('spells a name a bracketed #Const declares without its brackets, and leaves the brackets', () => {
		// The VBE also shows the declaration as `#Const BrName = 1`; the
		// formatter never changes a token's text.
		const project: Project = { modules: [['#Const [BrName] = 1', '#If brname Then', '#End If'], ['#If BRNAME Then', '#End If']] };
		expect(format(project, 0)).toEqual(['#Const [BrName] = 1', '#If BrName Then', '#End If']);
		expect(format(project, 1)).toEqual(['#If BrName Then', '#End If']);
	});

	it('cases directive lines in a branch that is not compiled, and keeps a trailing comment and a continuation', () => {
		expect(format(one(
			'#If False Then',
			'#if vba7 then',
			'#elseif win64 then',
			'#end if',
			'#End If',
			"#If vba7 Then ' trailing comment vba7",
			'#End If',
			'#If vba7 _',
			'    And win64 Then',
			'#End If',
		))).toEqual([
			'#If False Then',
			'    #If VBA7 Then',
			'    #ElseIf Win64 Then',
			'    #End If',
			'#End If',
			"#If VBA7 Then ' trailing comment vba7",
			'#End If',
			'#If VBA7 _',
			'    And Win64 Then',
			'#End If',
		]);
	});
});

describe('conditionalConstantDeclarations', () => {
	it('lists the names in source order, and not those a comment holds, reserved words or an unclosed bracket', () => {
		expect(conditionalConstantDeclarations([
			'#Const B = 1',
			"' #Const InComment = 1",
			"' note _",
			'#Const ContinuedComment = 1',
			'#If VBA7 Then',
			'    #Const A = 1',
			'#End If',
			'#Const Debug = 1',
			'#Const [Unclosed = 1',
			'Const Ordinary = 1',
		].join('\n'))).toEqual(['B', 'A']);
	});

	it('does not tokenize a module with no #Const', () => {
		const before = conditionalDeclarationScanCountForTests();
		expect(conditionalConstantDeclarations('Sub T()\n    Const X = 1 \' #If\nEnd Sub\n')).toEqual([]);
		expect(conditionalDeclarationScanCountForTests()).toBe(before);
		expect(conditionalConstantDeclarations('#Const X = 1\n')).toEqual(['X']);
		expect(conditionalDeclarationScanCountForTests()).toBe(before + 1);
	});
});
