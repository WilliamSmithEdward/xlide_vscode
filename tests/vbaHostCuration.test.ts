import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { getExcelObjectModel, type HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import {
	getHostMembers,
	resolveHostConstant,
	resolveHostEnum,
	resolveMemberReturnType,
} from '../src/analyzer/host/hostModel';
import { derivedConstantDoc, derivedMemberDoc } from '../src/analyzer/host/hostMemberDocs';
import {
	analyzeModule,
	resolveHover,
	resolveMemberCompletions,
	resolveTypeCompletions,
} from '../src/analyzer';

// The type libraries are faithful but lossy: an accessor that hands
// back a real object is declared `As Object`, a property's declared type never
// reaches a tooltip, and 9,555 enum-member descriptions were being read from the
// corpus and thrown away. These pin the curation that fixes each of those.

const MODELS: ReadonlyArray<[string, () => HostObjectModel]> = [
	['Excel', getExcelObjectModel],
	['Word', getWordObjectModel],
	['PowerPoint', getPowerPointObjectModel],
	['Access', getAccessObjectModel],
];

/** The reference corpus is developer-local (gitignored); corpus tests skip without it. */
const CORPUS = path.join(process.cwd(), 'reference', 'excel', 'json');
const hasCorpus = fs.existsSync(CORPUS);

describe('generic Object returns are repaired', () => {
	it('reproduces every return type Excel used to transcribe by hand', async () => {
		if (!hasCorpus) {
			return;
		}
		// Excel's model named these one member at a time. The curator derives them
		// from the corpus instead; disagreeing with a transcribed answer is the
		// failure this guards, and the count guards silent shrinkage of the rules.
		const { readDumps, createCurator, classNamesIn, CLASS_KINDS } =
			await import('../scripts/reference-curation.mjs');
		const dumps = readDumps(CORPUS);
		const curator = createCurator({
			dumps,
			prefix: 'Excel',
			foreignClasses: new Map(
				[...classNamesIn(path.join(process.cwd(), 'reference', 'office', 'json'))]
					.map((name: string) => [name, 'Office']),
			),
		});
		const model = getExcelObjectModel();
		let agreed = 0;
		const disagreements: string[] = [];
		for (const [owner, dump] of dumps as Map<string, Record<string, unknown>>) {
			if (!CLASS_KINDS.has(dump.kind as string)) {
				continue;
			}
			const modelled = model.types[`Excel.${owner}`];
			for (const [list, kind] of [
				[(dump.properties ?? []) as Record<string, string>[], 'property'],
				[(dump.methods ?? []) as Record<string, string>[], 'method'],
			] as const) {
				for (const raw of list) {
					if (raw.name.startsWith('_')) {
						continue;
					}
					if ((kind === 'property' ? raw.type : raw.returns) !== 'Object') {
						continue;
					}
					const transcribed = modelled?.members.find(
						(member) => member.name.toLowerCase() === raw.name.toLowerCase(),
					)?.returns;
					if (!transcribed?.startsWith('Excel.')) {
						continue;
					}
					const derived = curator.resolveReturn(owner, raw, kind).returns;
					if (derived === transcribed) {
						agreed += 1;
					} else {
						disagreements.push(`${owner}.${raw.name}: transcribed ${transcribed}, derived ${derived}`);
					}
				}
			}
		}
		expect(disagreements).toEqual([]);
		expect(agreed).toBeGreaterThanOrEqual(85);
	});

	/**
	 * Members Excel's model names that the reference dump does not list, each
	 * confirmed present in the installed Excel 16.0 type library (1,036
	 * typeinfos, read with LoadTypeLib) and in the Excel interop reference.
	 *
	 * They are missing from the dump, not from Excel: the introspection reads the
	 * public `Worksheet`/`Chart` interfaces, while these live on the hidden
	 * `_Worksheet` and `_Chart` dispatch interfaces the objects actually
	 * implement. `ActiveSheet.Buttons` and `ChartArea.Interior` are everyday
	 * legacy VBA; `PivotFilters.Add` and `SlicerCaches.Add` have published
	 * reference pages.
	 */
	const HIDDEN_INTERFACE_MEMBERS = new Set([
		// _Worksheet: the Excel 5.0 sheet-control accessors.
		'Excel.Worksheet.Arcs', 'Excel.Worksheet.Buttons', 'Excel.Worksheet.CheckBoxes',
		'Excel.Worksheet.DrawingObjects', 'Excel.Worksheet.Drawings', 'Excel.Worksheet.DropDowns',
		'Excel.Worksheet.GroupBoxes', 'Excel.Worksheet.Labels', 'Excel.Worksheet.Lines',
		'Excel.Worksheet.ListBoxes', 'Excel.Worksheet.OptionButtons', 'Excel.Worksheet.Ovals',
		'Excel.Worksheet.Pictures', 'Excel.Worksheet.Rectangles', 'Excel.Worksheet.ScrollBars',
		'Excel.Worksheet.Spinners', 'Excel.Worksheet.TextBoxes',
		// _Chart, marked "Reserved for internal use" but used by the reference's
		// own example: `With ActiveChart / .Type = xlLine`.
		'Excel.Chart.Type',
		// Documented on learn.microsoft.com, absent from the dump.
		'Excel.PivotFilters.Add', 'Excel.SlicerCaches.Add',
		// Pre-2007 chart formatting, still on the interfaces.
		'Excel.ChartArea.Border', 'Excel.ChartArea.Font', 'Excel.ChartArea.Interior',
		'Excel.PlotArea.Border', 'Excel.PlotArea.Interior',
		'Excel.Legend.Border', 'Excel.Legend.Font', 'Excel.Legend.Interior',
		'Excel.LegendKey.Border', 'Excel.LegendKey.Interior',
		'Excel.Walls.Border', 'Excel.Walls.Interior',
		'Excel.Floor.Border', 'Excel.Floor.Interior',
		'Excel.DownBars.Border', 'Excel.DownBars.Interior',
		'Excel.UpBars.Border', 'Excel.UpBars.Interior',
	]);

	it('names no member the Excel type library does not carry', async () => {
		if (!hasCorpus) {
			return;
		}
		// Excel's model is the one host model with hand-written member lists, and
		// entries in it had been copied onto the wrong type: TextFrame.HasText is a
		// TextFrame2 member, WebOptions.AlwaysSaveInDefaultEncoding belongs to
		// DefaultWebOptions, PivotTable.PivotFilters to PivotField,
		// Application.IconSets to Workbook. TimelineState.SetFilterDateRange2
		// existed nowhere at all. Completion offered every one of them. The dump is
		// introspected from the type library, so it decides what a type has - and
		// anything it omits has to be justified in the set above.
		const { readDumps } = await import('../scripts/reference-curation.mjs');
		const dumps = readDumps(CORPUS) as Map<string, Record<string, Array<{ name: string }>>>;
		const model = getExcelObjectModel();
		const invented: string[] = [];
		for (const [qualified, type] of Object.entries(model.types)) {
			if (!qualified.startsWith('Excel.')) {
				continue;
			}
			const dump = dumps.get(qualified.slice('Excel.'.length));
			if (!dump) {
				continue;   // no dump for this type: nothing to check it against
			}
			const known = new Set(
				[...(dump.properties ?? []), ...(dump.methods ?? []), ...(dump.events ?? [])]
					.map((member) => member.name.toLowerCase()),
			);
			for (const member of type.members) {
				const qualifiedMember = `${qualified}.${member.name}`;
				if (!known.has(member.name.toLowerCase()) && !HIDDEN_INTERFACE_MEMBERS.has(qualifiedMember)) {
					invented.push(qualifiedMember);
				}
			}
		}
		expect(invented).toEqual([]);
	});

	it('chains through Chart.Axes in every host that carries a chart', () => {
		// The measured failure: `Chart.Axes` is declared `As Object` in all three
		// libraries, so `.Chart.Axes(xlCategory).HasTitle` died at the Axes hop
		// everywhere except Excel, whose model had transcribed the return by hand.
		// Excel, Word and PowerPoint re-export the same chart object model.
		// Access's `Chart` is a report control of the same name, not this object.
		for (const [host, getModel] of MODELS.filter(([name]) => name !== 'Access')) {
			const model = getModel();
			const axes = resolveMemberReturnType(`${host}.Chart`, 'Axes', model);
			expect(axes, `${host}.Chart.Axes`).toBe(`${host}.Axes`);
			expect(resolveMemberReturnType(axes!, 'Item', model), `${host}.Axes.Item`).toBe(`${host}.Axis`);
		}
	});

	it('resolves the documented PowerPoint chart sample end to end', () => {
		// The sample on learn.microsoft.com/office/vba/api/powerpoint.axis.
		const src = [
			'Option Explicit',
			'Public Sub Test()',
			'    With ActivePresentation.Slides(1).Shapes(1)',
			'        If .HasChart Then',
			'            With .Chart.Axes.Item(xlCategory)',
			'                .HasTitle = True',
			'                .AxisTitle.Caption = "1994"',
			'            End With',
			'        End If',
			'    End With',
			'End Sub',
		].join('\n');
		const model = getPowerPointObjectModel();
		const signatureAt = (marker: string): string | undefined =>
			resolveHover(src, src.indexOf(marker) + 2, { model })?.signature;
		expect(signatureAt('.HasChart')).toBe('Shape.HasChart As MsoTriState');
		expect(signatureAt('.Chart')).toBe('Shape.Chart As Chart');
		expect(signatureAt('.HasTitle')).toBe('Axis.HasTitle As Boolean');
		expect(signatureAt('.AxisTitle')).toBe('Axis.AxisTitle As AxisTitle');
		expect(signatureAt('.Caption')).toBe('AxisTitle.Caption As String');
	});

	it('keeps a documented either/or ambiguous rather than picking one', () => {
		const word = getWordObjectModel();
		const context = getHostMembers('Word.KeyBinding', word).find((m) => m.name === 'Context');
		expect(context?.returns).toBeUndefined();
		expect(context?.returnsAnyOf).toEqual(['Word.Document', 'Word.Template', 'Word.Application']);
	});

	it('never names a primitive as a chaining return', () => {
		// `returns` is the chain type: an absent one means "this hop states no
		// object". A bare `Long` there would be read as a type to chain into.
		for (const [host, getModel] of MODELS) {
			const offenders: string[] = [];
			for (const [qualified, type] of Object.entries(getModel().types)) {
				for (const member of type.members) {
					if (member.returns && !member.returns.includes('.')) {
						offenders.push(`${qualified}.${member.name} -> ${member.returns}`);
					}
				}
			}
			expect(offenders.slice(0, 5), host).toEqual([]);
		}
	});
});

describe('the shared Office library binds to its host', () => {
	it('resolves an Office object back to the host application', () => {
		// The Office library has no Application class - the host supplies it - so
		// every Application member across its types is declared `As Object`.
		for (const [host, getModel] of MODELS) {
			expect(
				resolveMemberReturnType('Office.TextRange2', 'Application', getModel()),
				host,
			).toBe(`${host}.Application`);
		}
	});

	it('gives each host its own binding of the same shared type', () => {
		expect(resolveMemberReturnType('Office.ColorFormat', 'Application', getWordObjectModel()))
			.toBe('Word.Application');
		expect(resolveMemberReturnType('Office.ColorFormat', 'Application', getExcelObjectModel()))
			.toBe('Excel.Application');
	});
});

describe('every member says what it is', () => {
	it.each(MODELS)('%s states a declared type or a call signature for every member', (host, getModel) => {
		const silent: string[] = [];
		for (const [qualified, type] of Object.entries(getModel().types)) {
			for (const member of type.members) {
				if (member.signature || member.declaredType || member.returns || member.returnsAnyOf) {
					continue;
				}
				silent.push(`${qualified}.${member.name}`);
			}
		}
		expect(silent, host).toEqual([]);
	});

	it.each(MODELS)('%s describes every member, from the reference or from the declaration', (host, getModel) => {
		const model = getModel();
		let documented = 0;
		let derived = 0;
		for (const [qualified, type] of Object.entries(model.types)) {
			for (const member of type.members) {
				if (member.doc?.summary) {
					documented += 1;
					continue;
				}
				const composed = derivedMemberDoc(member, qualified);
				expect(composed?.summary, `${qualified}.${member.name}`).toBeTruthy();
				expect(composed?.source).toBe('derived');
				derived += 1;
			}
		}
		expect(documented, host).toBeGreaterThan(1000);
		expect(documented + derived, host).toBeGreaterThan(4000);
	});

	it('composes a description only from what the type library declares', () => {
		const doc = derivedMemberDoc(
			{ name: 'Row', kind: 'property', declaredType: 'Long', access: 'read-only' },
			'Excel.Range',
		);
		expect(doc?.summary).toBe('Read-only Long property of the Range object.');
		expect(doc?.source).toBe('derived');
		expect(
			derivedMemberDoc({ name: 'Activate', kind: 'method', signature: 'Activate() As Variant' }, 'Excel.Worksheet')
				?.summary,
		).toBe('Method of the Worksheet object. Returns Variant.');
	});

	it('leaves a documented member alone', () => {
		const member = getHostMembers('Excel.Range', getExcelObjectModel()).find((m) => m.name === 'Value');
		expect(member?.doc?.summary).toContain('value of the specified range');
		expect(derivedMemberDoc(member!, 'Excel.Range')).toBeUndefined();
	});
});

describe('hover shows the curated facts', () => {
	const src = 'Sub T()\n    ActiveSheet.Range("A1").Value = 1\n    Dim n As Long\n    n = ActiveSheet.Range("A1").Row\nEnd Sub\n';

	it('renders a property with its declared type and read/write contract', () => {
		const value = resolveHover(src, src.indexOf('.Value') + 2, {});
		expect(value?.signature).toBe('Range.Value As Variant');
		expect(value?.details).toEqual(['Excel host property (read/write)']);

		const row = resolveHover(src, src.indexOf('.Row') + 2, {});
		expect(row?.signature).toBe('Range.Row As Long');
		expect(row?.details).toEqual(['Excel host property (read-only)']);
	});

	it('marks a composed description as derived so it never reads as transcribed', () => {
		const undocumented = 'Sub T()\n    Dim d As Diagram\n    d.AutoLayout = 1\nEnd Sub\n';
		const hover = resolveHover(undocumented, undocumented.indexOf('.AutoLayout') + 2, {});
		expect(hover?.signature).toContain('AutoLayout');
		expect(hover?.details.some((line) => line.includes('derived'))).toBe(true);
	});

	it('describes an enum constant from the reference', () => {
		const constant = resolveHostConstant('xlCategory', getPowerPointObjectModel());
		expect(constant?.doc?.summary).toBe('Axis displays categories.');
		const src2 = 'Sub T()\n    Dim a As Long\n    a = xlCategory\nEnd Sub\n';
		const hover = resolveHover(src2, src2.indexOf('xlCategory') + 2, {
			model: getPowerPointObjectModel(),
		});
		expect(hover?.documentation).toBe('Axis displays categories.');
	});

	it('describes a host type from the reference', () => {
		const src2 = 'Sub T()\n    Dim s As InlineShape\nEnd Sub\n';
		const hover = resolveHover(src2, src2.indexOf('InlineShape') + 2, { model: getWordObjectModel() });
		expect(hover?.details).toEqual(['Word host type']);
		expect(hover?.documentation).toContain('text layer of a document');
	});
});

describe('enum constants carry the reference description', () => {
	it.each(MODELS)('%s describes most of its enum members', (host, getModel) => {
		const constants = Object.values(getModel().constants ?? {});
		const described = constants.filter((constant) => constant.doc?.summary).length;
		expect(constants.length, host).toBeGreaterThan(500);
		expect(described / constants.length, host).toBeGreaterThan(0.5);
	});
});

describe('enumerations are part of the model', () => {
	it.each(MODELS)('%s resolves every enumeration it declares as a declared type', (host, getModel) => {
		const model = getModel();
		const names = Object.keys(model.enums ?? {});
		expect(names.length, host).toBeGreaterThan(90);
		// `Dim k As XlAxisType` is ordinary VBA; before this the name resolved to
		// nothing at all - no completion, no hover, no coloring.
		for (const name of names) {
			expect(resolveHostEnum(name, model)?.displayName, `${host}.${name}`).toBe(name);
		}
	});

	it('offers an enumeration in a declaration type position, as an enum', () => {
		const src = 'Sub T()\n    Dim k As XlAxis\nEnd Sub\n';
		const offered = resolveTypeCompletions(src, src.indexOf('XlAxis') + 6, {})
			.map((item) => `${item.name}:${item.kind}`);
		expect(offered).toContain('XlAxisType:enum');
	});

	it('describes an enumeration on hover', () => {
		const src = 'Sub T()\n    Dim k As XlAxisType\nEnd Sub\n';
		const hover = resolveHover(src, src.indexOf('XlAxisType') + 2, {});
		expect(hover?.signature).toBe('Enum XlAxisType');
		expect(hover?.details).toEqual(['Excel enum']);
		expect(hover?.documentation).toBe('Specifies the axis type.');
	});

	it('reaches an enumeration constants through the enum name', () => {
		const src = 'Sub T()\n    k = XlAxisType.\nEnd Sub\n';
		const members = resolveMemberCompletions(src, src.indexOf('XlAxisType.') + 11, {} as never)
			.map((item) => item.name);
		expect(members).toEqual(['xlCategory', 'xlSeriesAxis', 'xlValue']);
	});

	it('does not call an enum qualifier an undeclared variable', () => {
		const src = 'Option Explicit\nSub T()\n    Dim k As Long\n    k = XlAxisType.xlCategory\nEnd Sub\n';
		expect(analyzeModule(src, { host: 'excel', knownIdentifiers: new Set<string>() })).toEqual([]);
	});

	it.each(MODELS)('%s describes every enum constant, from the reference or its enumeration', (host, getModel) => {
		const model = getModel();
		const constants = Object.values(model.constants ?? {});
		expect(constants.length, host).toBeGreaterThan(500);
		for (const constant of constants) {
			const described = constant.doc?.summary
				?? derivedConstantDoc(constant, constant.type ? resolveHostEnum(constant.type, model) : undefined)?.summary;
			expect(described, `${host}: ${constant.name}`).toBeTruthy();
		}
	});
});

describe('a shared page does not name the wrong application', () => {
	// The reference publishes one page per shared object under every host
	// namespace without substituting the application name, so a PowerPoint
	// developer hovering Axis.ReversePlotOrder read "True if Microsoft Word
	// plots data points from last to first".
	it('names PowerPoint as the actor in PowerPoint', () => {
		const member = getHostMembers('PowerPoint.Axis', getPowerPointObjectModel())
			.find((item) => item.name === 'ReversePlotOrder');
		expect(member?.doc?.summary).toBe(
			'True if Microsoft PowerPoint plots data points from last to first. Read/write Boolean.',
		);
	});

	it('leaves a sentence written about another application alone', () => {
		// A chart's data really does live in an Excel workbook, whichever host
		// reads it, so this mention is the reference meaning what it says.
		const member = getHostMembers('PowerPoint.ChartData', getPowerPointObjectModel())
			.find((item) => item.name === 'IsLinked');
		expect(member?.doc?.summary).toContain('Microsoft Excel workbook');
	});

	it('leaves no foreign actor anywhere in the PowerPoint model', () => {
		const offenders: string[] = [];
		for (const [qualified, type] of Object.entries(getPowerPointObjectModel().types)) {
			if (!qualified.startsWith('PowerPoint.')) {
				continue;
			}
			for (const member of type.members) {
				if (member.doc?.summary?.includes('Microsoft Word')) {
					offenders.push(`${qualified}.${member.name}`);
				}
			}
		}
		expect(offenders).toEqual([]);
	});
});
