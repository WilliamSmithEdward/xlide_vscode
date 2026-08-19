import { describe, it, expect } from 'vitest';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import {
	collectHostGlobalTokens,
	collectHostMemberMethodTokens,
	ProjectIndex,
	resolveTypeSemanticTokens,
	type HostMemberTokenContext,
	type TypeCompletionContext,
	type VbaProjectTypeName,
} from '../src/analyzer';

function tokenTexts(
	source: string,
	ctx: TypeCompletionContext = {},
): Array<{ text: string; type: string }> {
	return resolveTypeSemanticTokens(source, ctx).map((t) => ({
		text: source.slice(t.span.start, t.span.end),
		type: t.tokenType,
	}));
}

describe('project type semantic tokens', () => {
	it('marks resolved class module names in As type positions', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As Person\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Person', type: 'class' },
		]);
	});

	it('marks UDTs, enums, parameters, returns, fields, and nested locals', () => {
		const source = [
			'Public Type Wrapper',
			'    Item As Person',
			'End Type',
			'Public Function Make(ByVal mode As Color) As Person',
			'    If True Then',
			'        Dim inner As Wrapper',
			'    End If',
			'End Function',
		].join('\n');
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
			{ name: 'Color', kind: 'enum', moduleName: 'Types' },
			{ name: 'Wrapper', kind: 'userType', moduleName: 'Types' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'Color', type: 'enum' },
			{ text: 'Person', type: 'class' },
			{ text: 'Wrapper', type: 'struct' },
		]);
	});

	it('does not mark unresolved names or non-type positions', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim value As Missing\n' +
			'    Person = 1\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([]);
	});

	it('marks project classes in New expressions', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New person\n' +
			'    If TypeOf p Is Person Then Debug.Print "ok"\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'person', type: 'class' },
			{ text: 'Person', type: 'class' },
		]);
	});

	it('marks project classes in Implements statements', () => {
		const source = 'Implements Person\nImplements Excel.Worksheet\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Person', type: 'class' },
		]);
	});

	it('uses the project binder visible type names as input', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim customer As Person\n' +
			'    Dim state As Status\n' +
			'End Sub\n';
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source,
		});
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: '',
		});
		index.setModule({
			moduleName: 'SharedTypes',
			moduleKind: 'standard',
			source: 'Public Enum Status\n    Active\nEnd Enum\n',
		});

		expect(tokenTexts(source, { projectTypes: index.visibleTypeNames('Caller') })).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'Status', type: 'enum' },
		]);
	});

	it('marks qualified project type names in As, New, and TypeOf positions', () => {
		const source = [
			'Public Sub T()',
			'    Dim point As Geometry.TPoint',
			'    Dim state As Workflow.Status',
			'    Set point = New Geometry.TPoint',
			'    If TypeOf point Is Geometry.TPoint Then Debug.Print "ok"',
			'End Sub',
		].join('\n');
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'TPoint', kind: 'class', moduleName: 'Geometry' },
			{ name: 'Status', kind: 'enum', moduleName: 'Workflow' },
			{ name: 'TPoint', kind: 'class', moduleName: 'OtherGeometry' },
		];

		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'TPoint', type: 'class' },
			{ text: 'Status', type: 'enum' },
			{ text: 'TPoint', type: 'class' },
			{ text: 'TPoint', type: 'class' },
		]);
	});
});

describe('type semantic tokens', () => {
	it('marks primitive, host, and project types with distinct token categories', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim amount As Currency\n' +
			'    Dim p As Person\n' +
			'    Dim ws As Worksheet\n' +
			'    Dim state As Status\n' +
			'    Dim point As TPoint\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
			{ name: 'Status', kind: 'enum', moduleName: 'Types' },
			{ name: 'TPoint', kind: 'userType', moduleName: 'Types' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Currency', type: 'type' },
			{ text: 'Person', type: 'class' },
			{ text: 'Worksheet', type: 'class' },
			{ text: 'Status', type: 'enum' },
			{ text: 'TPoint', type: 'struct' },
		]);
	});

	it('lets project types shadow primitive names for coloring', () => {
		const source = 'Public Sub T()\n    Dim value As Long\nEnd Sub\n';
		expect(
			tokenTexts(source, {
				projectTypes: [{ name: 'Long', kind: 'class', moduleName: 'Long' }],
			}),
		).toEqual([{ text: 'Long', type: 'class' }]);
	});

	it('colors colliding project type names generically', () => {
		const source = 'Public Sub T()\n    Dim value As Status\nEnd Sub\n';
		expect(
			tokenTexts(source, {
				projectTypes: [
					{ name: 'Status', kind: 'class', moduleName: 'StatusClass' },
					{ name: 'Status', kind: 'enum', moduleName: 'SharedTypes' },
				],
			}),
		).toEqual([{ text: 'Status', type: 'type' }]);
	});
});

describe('host-global semantic tokens', () => {
	function hostTokens(source: string): Array<{ text: string; type: string; modifiers: string[] }> {
		return collectHostGlobalTokens(source).map((t) => ({
			text: source.slice(t.span.start, t.span.end),
			type: t.tokenType,
			modifiers: t.modifiers ?? [],
		}));
	}

	it('marks a host global used as a receiver, not its member', () => {
		const source = 'Public Sub T()\n    Set wb = Application.ActiveWorkbook\nEnd Sub\n';
		expect(hostTokens(source)).toEqual([
			{ text: 'Application', type: 'variable', modifiers: ['defaultLibrary'] },
		]);
	});

	it('marks ThisWorkbook and a bare host-global value', () => {
		expect(hostTokens('Sub T()\n    Set s = ThisWorkbook.Sheets\nEnd Sub\n').map((t) => t.text))
			.toEqual(['ThisWorkbook']);
		expect(hostTokens('Sub T()\n    MsgBox Application.Name\nEnd Sub\n').map((t) => t.text))
			.toEqual(['Application']);
	});

	it('does not mark a host name in a type position (owned by the type collector)', () => {
		expect(hostTokens('Sub T()\n    Dim app As Application\nEnd Sub\n')).toEqual([]);
	});

	it('does not mark a member access on another receiver', () => {
		// `Range` is the host-global receiver; `.Application` is a member.
		expect(hostTokens('Sub T()\n    x = Range("A1").Application\nEnd Sub\n').map((t) => t.text))
			.toEqual(['Range']);
	});

	it('does not mark a name shadowed by a local or parameter', () => {
		expect(hostTokens('Sub T()\n    Dim Application As Long\n    Application = 1\nEnd Sub\n')).toEqual([]);
		expect(hostTokens('Sub T(ByVal Application As Long)\n    x = Application\nEnd Sub\n')).toEqual([]);
	});

	it('does not mark a name shadowed by a designer-declared control (issue #30)', () => {
		// Inside the form, the control wins name binding; the receiver must not
		// wear the host global's tint while the member lookup treats it as the
		// control (#29's shadow rule, extended to the global collector).
		const source = 'Sub T()\n    ActiveSheet.Clear\n    x = ActiveCell.Value\nEnd Sub\n';
		const controls = [{ name: 'ActiveSheet', type: 'MSForms.ListBox' }];
		const names = collectHostGlobalTokens(source, undefined, controls)
			.map((t) => source.slice(t.span.start, t.span.end));
		expect(names).not.toContain('ActiveSheet');
		expect(names).toContain('ActiveCell');
	});
});

describe('host-global tokens follow the module host (issue #24)', () => {
    it('paints Word globals under the Word model and never Excel globals', () => {
        const source = [
            'Sub T()',
            '    ActiveDocument.Save',
            '    ActiveSheet.Calculate',
            'End Sub',
            '',
        ].join('\r\n');
        const wordNames = collectHostGlobalTokens(source, getWordObjectModel())
            .map((token) => source.slice(token.span.start, token.span.end));
        expect(wordNames).toContain('ActiveDocument');
        expect(wordNames).not.toContain('ActiveSheet');

        const excelNames = collectHostGlobalTokens(source)
            .map((token) => source.slice(token.span.start, token.span.end));
        expect(excelNames).toContain('ActiveSheet');
        expect(excelNames).not.toContain('ActiveDocument');
    });
});

describe('host member method semantic tokens (issue #29)', () => {
	function methodTokens(source: string, ctx: HostMemberTokenContext = {}): string[] {
		return collectHostMemberMethodTokens(source, ctx).map(
			(t) => `${source.slice(t.span.start, t.span.end)}:${t.tokenType}`,
		);
	}

	it('paints resolved host method calls and leaves properties alone', () => {
		const source = [
			'Sub T()',
			'    ActiveSheet.Calculate',
			'    Application.Quit',
			'    x = ActiveSheet.Name',
			'End Sub',
			'',
		].join('\n');
		expect(methodTokens(source)).toEqual(['Calculate:function', 'Quit:function']);
	});

	it('never paints an unresolved member', () => {
		expect(methodTokens('Sub T()\n    ActiveSheet.NotAMember\nEnd Sub\n')).toEqual([]);
	});

	it('does not paint when a declaration shadows the receiver', () => {
		const source =
			'Sub T()\n    Dim ActiveSheet As Long\n    ActiveSheet.Calculate\nEnd Sub\n';
		expect(methodTokens(source)).toEqual([]);
	});

	it('leaves longer chains and With-block members out of scope', () => {
		const chain = 'Sub T()\n    ActiveSheet.Range("A1").Calculate\nEnd Sub\n';
		expect(methodTokens(chain)).toEqual([]);
		const withBlock =
			'Sub T()\n    With ActiveSheet\n        .Calculate\n    End With\nEnd Sub\n';
		expect(methodTokens(withBlock)).toEqual([]);
	});

	it('follows the module host: Word paints FitToPages, Excel does not', () => {
		const source = 'Sub T()\n    ActiveDocument.FitToPages\nEnd Sub\n';
		expect(methodTokens(source, { model: getWordObjectModel() }))
			.toEqual(['FitToPages:function']);
		expect(methodTokens(source)).toEqual([]);
	});

	it('resolves code-name receivers to their document host type', () => {
		const excel = 'Sub T()\n    Sheet1.Calculate\nEnd Sub\n';
		expect(methodTokens(excel, { codeNames: { sheet1: 'Excel.Worksheet' } }))
			.toEqual(['Calculate:function']);
		expect(methodTokens(excel)).toEqual([]);

		const word = 'Sub T()\n    ThisDocument.Save\nEnd Sub\n';
		expect(
			methodTokens(word, {
				model: getWordObjectModel(),
				codeNames: { thisdocument: 'Word.Document' },
			}),
		).toEqual(['Save:function']);
	});

	it('lets a designer-declared control shadow a host-global receiver', () => {
		const source = 'Sub T()\n    ActiveSheet.Calculate\nEnd Sub\n';
		expect(
			methodTokens(source, {
				implicitMembers: [{ name: 'ActiveSheet', type: 'MSForms.ListBox' }],
			}),
		).toEqual([]);
	});

	it('paints Me-qualified host methods in a document module (issue #31)', () => {
		const source = 'Sub T()\n    Me.Calculate\n    x = Me.Name\nEnd Sub\n';
		expect(methodTokens(source, { meType: 'Excel.Worksheet' }))
			.toEqual(['Calculate:function']);
		// Without a Me host type - a standard module, a loose file - Me stays
		// plain, and an MSForms Me belongs to the implicit-member collector.
		expect(methodTokens(source)).toEqual([]);
		expect(methodTokens('Sub T()\n    Me.Hide\nEnd Sub\n', { meType: 'MSForms.UserForm' }))
			.toEqual([]);
	});

	it('paints methods on a declared local with a host type (issue #33)', () => {
		const word = [
			'Public Sub T()',
			'    Dim rng As Range',
			'    Set rng = ActiveDocument.Range(0, 0)',
			'    rng.InsertParagraphAfter',
			'End Sub',
			'',
		].join('\n');
		// ActiveDocument.Range already painted (#29); the local's member joins it.
		expect(methodTokens(word, { model: getWordObjectModel() }))
			.toEqual(['Range:function', 'InsertParagraphAfter:function']);

		const excel = 'Sub T()\n    Dim ws As Worksheet\n    ws.Calculate\n    x = ws.Name\nEnd Sub\n';
		expect(methodTokens(excel)).toEqual(['Calculate:function']);

		const param = 'Sub T(ByVal ws As Worksheet)\n    ws.Calculate\nEnd Sub\n';
		expect(methodTokens(param)).toEqual(['Calculate:function']);
	});

	it('refuses ambiguous, untyped, and project-shadowed locals (issue #33)', () => {
		const ambiguous = [
			'Sub A()',
			'    Dim rng As Range',
			'    rng.Calculate',
			'End Sub',
			'Sub B()',
			'    Dim rng As Long',
			'    rng.Calculate',
			'End Sub',
			'',
		].join('\n');
		expect(methodTokens(ambiguous)).toEqual([]);

		const untyped = 'Sub T()\n    Dim rng\n    rng.Calculate\nEnd Sub\n';
		expect(methodTokens(untyped)).toEqual([]);

		// A project class named Range wins the As clause; the host reading
		// must not paint over it.
		const shadowed = 'Sub T()\n    Dim rng As Range\n    rng.Calculate\nEnd Sub\n';
		expect(methodTokens(shadowed, { projectTypes: [{ name: 'Range', kind: 'class' }] }))
			.toEqual([]);
	});

	it('paints a bare host Global method and value (issue #34)', () => {
		const word = 'Sub T()\n    TopM = InchesToPoints(1)\n    Set x = RecentFiles\nEnd Sub\n';
		const wordTokens = collectHostGlobalTokens(word, getWordObjectModel()).map(
			(t) => `${word.slice(t.span.start, t.span.end)}:${t.tokenType}`,
		);
		expect(wordTokens).toContain('InchesToPoints:function');
		expect(wordTokens).toContain('RecentFiles:variable');

		const excel = 'Sub T()\n    Set r = Union(a, b)\nEnd Sub\n';
		const excelTokens = collectHostGlobalTokens(excel).map(
			(t) => `${excel.slice(t.span.start, t.span.end)}:${t.tokenType}`,
		);
		expect(excelTokens).toContain('Union:function');

		// Shadowing and member access keep their gates.
		const shadowed =
			'Sub T()\n    Dim InchesToPoints As Long\n    x = InchesToPoints\nEnd Sub\n';
		expect(collectHostGlobalTokens(shadowed, getWordObjectModel())).toEqual([]);
		const memberAccess = 'Sub T()\n    x = Foo.InchesToPoints(1)\nEnd Sub\n';
		expect(collectHostGlobalTokens(memberAccess, getWordObjectModel())).toEqual([]);
	});

	it('paints Word Me members and keeps chains and With out of scope', () => {
		expect(
			methodTokens('Sub T()\n    Me.FitToPages\nEnd Sub\n', {
				model: getWordObjectModel(),
				meType: 'Word.Document',
			}),
		).toEqual(['FitToPages:function']);
		const chain = 'Sub T()\n    Me.Range("A1").Calculate\nEnd Sub\n';
		expect(methodTokens(chain, { meType: 'Excel.Worksheet' })).toEqual([]);
		const withBlock = 'Sub T()\n    With Me\n        .Calculate\n    End With\nEnd Sub\n';
		expect(methodTokens(withBlock, { meType: 'Excel.Worksheet' })).toEqual([]);
	});
});
