// Which Word, PowerPoint and Access types can prove a member absent (issue
// #127). The closed sets come from the registered type libraries (Word 8.7,
// PowerPoint 2.12, Access 9.0, NONEXTENSIBLE on the coclass's default
// interface), and the three samples the issue measured through pyVBAharness
// on 2026-09-25 are the oracle: Word Range, PowerPoint Shape and an Access
// TextBox refuse a member they lack, while Access Control and Form compile
// whatever is named on them.

import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { hostTypeResolvesWhenCompiling } from '../src/analyzer/host/typeExtensibility';
import { byCode } from './helpers/diagnostics';

const found = (src: string, hostModel: ReturnType<typeof getWordObjectModel>): string[] =>
	byCode(analyzeModule(src, { hostModel, moduleKind: 'standard' }), 'member-not-found').map((one) => one.message);

describe('closed and open types per host', () => {
	it('reads the flags the type libraries carry', () => {
		for (const closed of ['Word.Range', 'Word.Selection', 'Word.Paragraph', 'Word.Table', 'PowerPoint.Slide', 'PowerPoint.Shape', 'PowerPoint.TextRange', 'Access.TextBox', 'Access.ComboBox']) {
			expect(hostTypeResolvesWhenCompiling(closed), closed).toBe(true);
		}
		for (const open of ['Word.Document', 'Access.Form', 'Access.Report', 'Access.Control', 'Access.CodeProject', 'Access.CurrentProject', 'Office.CommandBar']) {
			expect(hostTypeResolvesWhenCompiling(open), open).toBe(false);
		}
	});
});

describe('member-not-found outside Excel', () => {
	it('reports a member Word Range has not got, and leaves Document alone', () => {
		const model = getWordObjectModel();
		expect(found('Sub Main()\n    Dim r As Range\n    Set r = ActiveDocument.Content\n    r.SomeCustomMacro\nEnd Sub\n', model))
			.toEqual(["Method or data member not found: 'Word.Range.SomeCustomMacro'."]);
		expect(found('Sub Main()\n    ActiveDocument.SomeCustomMacro\n    ActiveDocument.Content.InsertAfter "x"\nEnd Sub\n', model)).toEqual([]);
	});

	it('reports a member a PowerPoint Shape has not got', () => {
		const model = getPowerPointObjectModel();
		expect(found('Sub Main()\n    Dim s As Shape\n    Set s = ActivePresentation.Slides(1).Shapes(1)\n    s.SomeCustomMacro\n    s.TextFrame.TextRange.Text = "x"\nEnd Sub\n', model))
			.toEqual(["Method or data member not found: 'PowerPoint.Shape.SomeCustomMacro'."]);
	});

	it('reports RowSource on an Access TextBox and allows it on Control and Form', () => {
		const model = getAccessObjectModel();
		expect(found('Option Compare Database\nSub Main()\n    Dim t As TextBox\n    Dim x As Variant\n    x = t.RowSource\nEnd Sub\n', model))
			.toHaveLength(1);
		expect(found('Option Compare Database\nSub Main()\n    Dim ctl As Control, f As Form\n    Dim x As Variant\n    x = ctl.RowSource\n    x = f.CustomerID\nEnd Sub\n', model)).toEqual([]);
	});
});
