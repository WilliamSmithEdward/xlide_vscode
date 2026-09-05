// Annotations that stand for a module's hidden attributes, parity with
// xlide_vbide. A VBA module carries attributes the code pane never shows and
// the editor gives no way to set; an annotation names one in the code, where it
// can be read and reviewed, and the rewriter writes the attribute to match.
//
// Placement is what binds an annotation. Getting that wrong is a problem
// reported with its line, never a guess about what was meant.

import { describe, expect, it } from 'vitest';
import {
	readAttributeAnnotations,
	type AnnotationKind,
} from '../src/analyzer/annotations/attributeAnnotations';
import { applyAttributeAnnotations } from '../src/analyzer/annotations/attributeRewriter';

const CLASS_HEADER = [
	'VERSION 1.0 CLASS',
	'BEGIN',
	'  MultiUse = -1  \'True',
	'END',
	'Attribute VB_Name = "CThing"',
	'Attribute VB_GlobalNameSpace = False',
	'Attribute VB_Creatable = False',
	'Attribute VB_PredeclaredId = False',
	'Attribute VB_Exposed = False',
];

const STANDARD_HEADER = ['Attribute VB_Name = "Module1"'];

function moduleOf(header: readonly string[], body: readonly string[]): string {
	return [...header, ...body].join('\r\n');
}

function kinds(source: string): Array<[AnnotationKind, string | undefined, string | undefined]> {
	return readAttributeAnnotations(source).annotations
		.map((one) => [one.kind, one.argument, one.target]);
}

function problems(source: string): string[] {
	return readAttributeAnnotations(source).problems.map((one) => one.message);
}

function rewritten(source: string): string {
	return applyAttributeAnnotations(source, readAttributeAnnotations(source)).text;
}

describe('reading an annotation', () => {
	it('accepts every spelling of the argument', () => {
		for (const line of [
			'\'@ModuleDescription("A thing")',
			'\'@ModuleDescription(A thing)',
			'\'@ModuleDescription "A thing"',
			'\'@moduledescription("A thing")',
		]) {
			const found = readAttributeAnnotations(`${line}\r\nOption Explicit\r\n`).annotations;
			expect(found.map((one) => one.kind), line).toEqual(['ModuleDescription']);
			expect(found[0].argument, line).toBe('A thing');
		}
	});

	it('reads a bare annotation with no argument', () => {
		expect(kinds('\'@PredeclaredId\r\nOption Explicit\r\n'))
			.toEqual([['PredeclaredId', undefined, undefined]]);
	});

	it('binds a member annotation to the procedure below it', () => {
		const source = [
			'Option Explicit',
			'',
			'\'@Description("Totals the rows")',
			'Public Function Total() As Long',
			'End Function',
		].join('\r\n');
		expect(kinds(source)).toEqual([['Description', 'Totals the rows', 'Total']]);
	});

	it('binds a variable description to the declaration below it', () => {
		const source = [
			'Option Explicit',
			'\'@VariableDescription("How many")',
			'Public Count As Long',
		].join('\r\n');
		expect(kinds(source)).toEqual([['VariableDescription', 'How many', 'Count']]);
	});

	it('leaves other people\'s annotations alone', () => {
		// '@Folder is the explorer's, and '@Ignore and '@TestMethod are
		// Rubberduck's own. None of them is an attribute.
		const source = [
			'\'@Folder("Accounts")',
			'\'@Ignore ProcedureNotUsed',
			'\'@TestMethod("Unit")',
			'Option Explicit',
		].join('\r\n');
		expect(readAttributeAnnotations(source)).toEqual({ annotations: [], problems: [] });
	});
});

describe('placement problems, reported rather than guessed', () => {
	it('refuses a module annotation below the first procedure', () => {
		const source = [
			'Option Explicit',
			'Public Sub Go()',
			'End Sub',
			'\'@PredeclaredId',
		].join('\r\n');
		expect(problems(source)).toEqual([
			"'@PredeclaredId is a module annotation and belongs in the declarations section, above the first procedure.",
		]);
	});

	it('refuses a member annotation above something that is not a procedure', () => {
		const source = ['Option Explicit', '\'@Description("x")', 'Public Total As Long'].join('\r\n');
		expect(problems(source)[0]).toMatch(/describes a procedure, and 'Total' is a variable/);
	});

	it('refuses a variable description above a procedure', () => {
		const source = ['Option Explicit', '\'@VariableDescription("x")', 'Public Sub Go()', 'End Sub'].join('\r\n');
		expect(problems(source)[0]).toMatch(/describes a module-level variable, and 'Go' is a procedure/);
	});

	it('refuses an annotation above nothing at all', () => {
		expect(problems('Option Explicit\r\n\'@Description("x")\r\n')[0])
			.toMatch(/is above nothing/);
	});

	it('refuses one that needs text and has none', () => {
		expect(problems('\'@ModuleDescription\r\nOption Explicit\r\n')[0])
			.toMatch(/needs the text to write, in brackets/);
	});

	it('refuses a second default member, naming the line that already has it', () => {
		const source = [
			'Option Explicit',
			'\'@DefaultMember',
			'Public Function A() As Long',
			'End Function',
			'\'@DefaultMember',
			'Public Function B() As Long',
			'End Function',
		].join('\r\n');
		expect(problems(source)[0]).toMatch(/appears again; a class has one default member, and line 2/);
	});

	it('refuses a hotkey that is not one letter', () => {
		const source = ['Option Explicit', '\'@ExcelHotkey("Ctrl")', 'Public Sub Go()', 'End Sub'].join('\r\n');
		expect(problems(source)[0]).toMatch(/takes one letter/);
	});

	it('keeps the first of a repeated module annotation and says so', () => {
		const source = ['\'@Exposed', '\'@Exposed', 'Option Explicit'].join('\r\n');
		expect(kinds(source)).toEqual([['Exposed', undefined, undefined]]);
		expect(problems(source)[0]).toMatch(/appears more than once; the first one counts/);
	});
});

describe('writing the attributes', () => {
	it('sets a class attribute the header already carries', () => {
		const source = moduleOf(CLASS_HEADER, ['\'@PredeclaredId', 'Option Explicit', '']);
		expect(rewritten(source)).toContain('Attribute VB_PredeclaredId = True');
		expect(rewritten(source)).toContain('Attribute VB_Exposed = False');
	});

	it('inserts a module description, which a header need not already have', () => {
		const source = moduleOf(STANDARD_HEADER, ['\'@ModuleDescription("Helpers")', 'Option Explicit', '']);
		expect(rewritten(source)).toContain('Attribute VB_Description = "Helpers"');
	});

	it('will not invent a class attribute on a standard module', () => {
		const source = moduleOf(STANDARD_HEADER, ['\'@PredeclaredId', 'Option Explicit', '']);
		const result = applyAttributeAnnotations(source, readAttributeAnnotations(source));
		expect(result.changes).toEqual([]);
		expect(result.skipped[0]).toMatch(/not an attribute this kind of module carries/);
	});

	it('writes a member attribute directly under its header', () => {
		const source = moduleOf(STANDARD_HEADER, [
			'Option Explicit',
			'',
			'\'@Description("Totals the rows")',
			'Public Function Total() As Long',
			'End Function',
			'',
		]);
		expect(rewritten(source)).toContain(
			'Public Function Total() As Long\r\nAttribute Total.VB_Description = "Totals the rows"',
		);
	});

	it('writes the numbers VBA wants for a default member and an enumerator', () => {
		const source = moduleOf(CLASS_HEADER, [
			'Option Explicit',
			'\'@DefaultMember',
			'Public Function Item() As Variant',
			'End Function',
			'\'@Enumerator',
			'Public Function NewEnum() As IUnknown',
			'End Function',
			'',
		]);
		const out = rewritten(source);
		expect(out).toContain('Attribute Item.VB_UserMemId = 0');
		expect(out).toContain('Attribute NewEnum.VB_UserMemId = -4');
	});

	it('writes a hotkey the way the editor stores it', () => {
		const source = moduleOf(STANDARD_HEADER, [
			'Option Explicit',
			'\'@ExcelHotkey("D")',
			'Public Sub Report()',
			'End Sub',
			'',
		]);
		expect(rewritten(source))
			.toContain('Attribute Report.VB_ProcData.VB_Invoke_Func = "D\\n14"');
	});

	it('describes a module-level variable', () => {
		const source = moduleOf(STANDARD_HEADER, [
			'Option Explicit',
			'\'@VariableDescription("How many")',
			'Public Count As Long',
			'',
		]);
		expect(rewritten(source)).toContain('Attribute Count.VB_VarDescription = "How many"');
	});

	it('doubles a quote inside the text, as VBA writes it', () => {
		const source = moduleOf(STANDARD_HEADER, ['\'@ModuleDescription("a ""quoted"" word")', 'Option Explicit', '']);
		expect(rewritten(source)).toContain('Attribute VB_Description = "a ""quoted"" word"');
	});

	it('updates an attribute that is already there rather than adding a second', () => {
		const source = moduleOf(STANDARD_HEADER, [
			'Option Explicit',
			'\'@Description("New words")',
			'Public Sub Go()',
			'Attribute Go.VB_Description = "Old words"',
			'End Sub',
			'',
		]);
		const out = rewritten(source);
		expect(out).toContain('Attribute Go.VB_Description = "New words"');
		expect(out).not.toContain('Old words');
		expect(out.match(/Go\.VB_Description/g)).toHaveLength(1);
	});

	it('reports what it changed, and from what', () => {
		const source = moduleOf(CLASS_HEADER, ['\'@PredeclaredId', 'Option Explicit', '']);
		const result = applyAttributeAnnotations(source, readAttributeAnnotations(source));
		expect(result.changes).toEqual([
			{ target: 'module', attribute: 'VB_PredeclaredId', from: 'False', to: 'True' },
		]);
	});

	it('changes nothing when the attributes already say what the annotations do', () => {
		const source = moduleOf(
			[...CLASS_HEADER.slice(0, 7), 'Attribute VB_PredeclaredId = True', 'Attribute VB_Exposed = False'],
			['\'@PredeclaredId', 'Option Explicit', ''],
		);
		const result = applyAttributeAnnotations(source, readAttributeAnnotations(source));
		expect(result.changes).toEqual([]);
		expect(result.text).toBe(source);
	});

	it('leaves every other line exactly as it was', () => {
		// The module is about to be written back over the developer's code.
		const body = [
			'Option Explicit',
			'',
			'\'@Description("Totals")',
			'Public Function Total() As Long',
			'    Total = 1  \' a trailing comment',
			'End Function',
			'',
		];
		const source = moduleOf(STANDARD_HEADER, body);
		const out = rewritten(source).split('\r\n');
		for (const line of [...STANDARD_HEADER, ...body]) {
			expect(out, line).toContain(line);
		}
		// Exactly one line added.
		expect(out).toHaveLength(STANDARD_HEADER.length + body.length + 1);
	});

	it('keeps the module\'s own line endings', () => {
		const source = moduleOf(STANDARD_HEADER, ['\'@ModuleDescription("x")', 'Option Explicit', '']).replace(/\r\n/g, '\n');
		expect(rewritten(source)).not.toContain('\r');
	});
});

describe('a property whose legs share a name', () => {
	it('puts each description on its own leg, however many lines are inserted above', () => {
		// The leg is counted when the annotation is READ. Working it out later
		// from a line number reads that number against text the rewriter has
		// already inserted into: with enough inserts above, the Get shifts past
		// the Let's recorded line and takes its description.
		const source = [
			'Attribute VB_Name = "CThing"',
			'\'@ModuleDescription("m")',
			'Option Explicit',
			'\'@Description("a")',
			'Public Sub A()',
			'End Sub',
			'\'@Description("b")',
			'Public Sub B()',
			'End Sub',
			'\'@Description("c")',
			'Public Sub C()',
			'End Sub',
			'\'@Description("reads it")',
			'Public Property Get Value() As Long',
			'End Property',
			'\'@Description("writes it")',
			'Public Property Let Value(ByVal v As Long)',
			'End Property',
			'',
		].join('\r\n');
		const lines = applyAttributeAnnotations(source, readAttributeAnnotations(source)).text.split('\r\n');
		const after = (needle: string): string => lines[lines.findIndex((l) => l.includes(needle)) + 1];
		expect(after('Property Get Value')).toBe('Attribute Value.VB_Description = "reads it"');
		expect(after('Property Let Value')).toBe('Attribute Value.VB_Description = "writes it"');
	});
});
