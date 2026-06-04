// Excel host object-model metadata for member completion.
//
// SOURCE OF TRUTH: the member names below are transcribed from the official
// Microsoft Office VBA object-model reference (learn.microsoft.com/office/vba/
// api/excel.*), verified on 2026-05-30. This is a curated subset of the most
// commonly used members of the core Excel automation types; it is NOT the full
// object model. Each type lists verified Properties and Methods. Return types
// are provided only for members whose result type is stable and documented, to
// enable member-access chaining (e.g. ws.Range(...).Offset(...).).
//
// This data is host metadata, NOT VBA language grammar. Per
// docs/xlide_vba_language_service_roadmap.md it must never override core
// MS-VBAL language resolution. LLM-generated member lists are never used here.

import {
	EXCEL_REFERENCE_ENUM_CONSTANTS,
	EXCEL_REFERENCE_MEMBER_SETS,
	EXCEL_WORKBOOK_REFERENCE_MEMBERS,
	EXCEL_WORKBOOK_REFERENCE_PROVENANCE,
} from './excelReferenceMembers';
import { OFFICE_REFERENCE_ENUM_CONSTANTS } from './officeReferenceConstants';
import type { VbaDoc } from '../docs/docModel';

// Events are retained as a distinct metadata kind for future handler-authoring
// surfaces, but object access (`Workbook.EventName`) exposes only properties and
// methods through hostModel.getHostMembers().
export type HostMemberKind = 'property' | 'method' | 'event';

export interface HostMember {
	name: string;
	kind: HostMemberKind;
	/** Qualified host type this member returns, when stable (for chaining). */
	returns?: string;
	/** Candidate host types this member may return, when the object model is mixed. */
	returnsAnyOf?: readonly string[];
	/** Verified call signature from reference metadata, when available. */
	signature?: string;
	/** Reference documentation rendered in completion, hover, and call tips. */
	doc?: VbaDoc;
}

export interface HostConstant {
	name: string;
	/** Host enum type this constant belongs to, e.g. "XlDirection". */
	type?: string;
	value?: string | number;
	source?: 'external' | 'verified';
	doc?: VbaDoc;
}

export interface HostType {
	/** Bare display name, e.g. "Worksheet". */
	displayName: string;
	/**
	 * True only when `members` is complete enough to prove a member is absent.
	 * Curated subsets must leave this false/undefined.
	 */
	exhaustive?: boolean;
	members: HostMember[];
}

export interface HostObjectModel {
	/** Provenance string shown in docs/verification map. */
	source: string;
	/** Qualified type name (e.g. "Excel.Range") -> type metadata. */
	types: Record<string, HostType>;
	/** Lowercased type name as written in `As <type>` -> qualified type. */
	aliases: Record<string, string>;
	/** Host-injected global identifier (canonical casing) -> qualified type. */
	globals: Record<string, string>;
	/** Host enum constants, keyed by canonical name. */
	constants?: Record<string, HostConstant>;
	/**
	 * Verified call signatures for callable members, keyed by qualified type
	 * then lowercased member name. Used by signature help (parameter info).
	 * Only members whose parameter list is transcribed from the Office VBA
	 * reference appear here; absence simply means no call tip is offered.
	 */
	memberSignatures?: Record<string, Record<string, string>>;
}

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const RANGE = 'Excel.Range';
const APPLICATION = 'Excel.Application';
const WORKBOOKS = 'Excel.Workbooks';
const WORKSHEETS = 'Excel.Worksheets';
const SHEETS = 'Excel.Sheets';
// Additional core Excel types, transcribed from the Excel COM type library and
// the Office VBA object-model reference to broaden member completion and enable
// deeper member-access chaining (e.g. Range.Font., ws.ListObjects(1).Range.).
const WINDOW = 'Excel.Window';
const WINDOWS = 'Excel.Windows';
const NAME = 'Excel.Name';
const NAMES = 'Excel.Names';
const COMMENT = 'Excel.Comment';
const COMMENTS = 'Excel.Comments';
const LISTOBJECT = 'Excel.ListObject';
const LISTOBJECTS = 'Excel.ListObjects';
const LISTROW = 'Excel.ListRow';
const LISTROWS = 'Excel.ListRows';
const LISTCOLUMN = 'Excel.ListColumn';
const LISTCOLUMNS = 'Excel.ListColumns';
const PIVOTTABLE = 'Excel.PivotTable';
const PIVOTTABLES = 'Excel.PivotTables';
const CHART = 'Excel.Chart';
const CHARTS = 'Excel.Charts';
const CHARTOBJECT = 'Excel.ChartObject';
const CHARTOBJECTS = 'Excel.ChartObjects';
const SHAPE = 'Excel.Shape';
const SHAPES = 'Excel.Shapes';
const FONT = 'Excel.Font';
const INTERIOR = 'Excel.Interior';
const BORDER = 'Excel.Border';
const BORDERS = 'Excel.Borders';
const AREAS = 'Excel.Areas';
const HYPERLINK = 'Excel.Hyperlink';
const HYPERLINKS = 'Excel.Hyperlinks';
const WORKSHEETFUNCTION = 'Excel.WorksheetFunction';
const STYLE = 'Excel.Style';
const STYLES = 'Excel.Styles';
const PAGESETUP = 'Excel.PageSetup';
const VALIDATION = 'Excel.Validation';

function p(name: string, returns?: string): HostMember {
	return { name, kind: 'property', returns };
}
function pAny(name: string, returnsAnyOf: readonly string[]): HostMember {
	return { name, kind: 'property', returnsAnyOf };
}
function m(name: string, returns?: string): HostMember {
	return { name, kind: 'method', returns };
}

function mergeHostMembers(
	primary: readonly HostMember[],
	reference: readonly HostMember[],
): HostMember[] {
	const out: HostMember[] = [];
	const byName = new Map<string, HostMember>();
	for (const member of [...primary, ...reference]) {
		const key = member.name.toLowerCase();
		const existing = byName.get(key);
		if (existing) {
			enrichHostMember(existing, member);
			continue;
		}
		const copy = { ...member };
		byName.set(key, copy);
		out.push(copy);
	}
	return out;
}

function enrichHostMember(target: HostMember, source: HostMember): void {
	target.returns ??= source.returns;
	target.returnsAnyOf ??= source.returnsAnyOf;
	target.signature ??= source.signature;
	target.doc ??= source.doc;
}

function referenceMembers(displayName: string): readonly HostMember[] {
	return EXCEL_REFERENCE_MEMBER_SETS[displayName] ?? [];
}

function mergeHostConstants(
	...sets: Array<Record<string, HostConstant>>
): Record<string, HostConstant> {
	const out: Record<string, HostConstant> = {};
	const keysByLowerName = new Map<string, string>();
	for (const set of sets) {
		for (const constant of Object.values(set)) {
			const lowerName = constant.name.toLowerCase();
			const previousKey = keysByLowerName.get(lowerName);
			if (previousKey) {
				delete out[previousKey];
			}
			keysByLowerName.set(lowerName, constant.name);
			out[constant.name] = constant;
		}
	}
	return out;
}

const EXCEL_HOST_ENUM_CONSTANTS = mergeHostConstants(
	OFFICE_REFERENCE_ENUM_CONSTANTS,
	EXCEL_REFERENCE_ENUM_CONSTANTS,
);

export const EXCEL_OBJECT_MODEL: HostObjectModel = {
	source: `Office VBA object-model reference (learn.microsoft.com) + Excel COM type library, verified 2026-05-30; promoted Excel and Office reference enum constants; promoted Excel reference metadata with Workbook/Worksheet exhaustive dumps; Workbook dump ${EXCEL_WORKBOOK_REFERENCE_PROVENANCE}`,
	aliases: {
		workbook: WORKBOOK,
		worksheet: WORKSHEET,
		range: RANGE,
		application: APPLICATION,
		workbooks: WORKBOOKS,
		worksheets: WORKSHEETS,
		sheets: SHEETS,
		window: WINDOW,
		windows: WINDOWS,
		name: NAME,
		names: NAMES,
		comment: COMMENT,
		comments: COMMENTS,
		listobject: LISTOBJECT,
		listobjects: LISTOBJECTS,
		listrow: LISTROW,
		listrows: LISTROWS,
		listcolumn: LISTCOLUMN,
		listcolumns: LISTCOLUMNS,
		pivottable: PIVOTTABLE,
		pivottables: PIVOTTABLES,
		chart: CHART,
		charts: CHARTS,
		chartobject: CHARTOBJECT,
		chartobjects: CHARTOBJECTS,
		shape: SHAPE,
		shapes: SHAPES,
		font: FONT,
		interior: INTERIOR,
		border: BORDER,
		borders: BORDERS,
		areas: AREAS,
		hyperlink: HYPERLINK,
		hyperlinks: HYPERLINKS,
		worksheetfunction: WORKSHEETFUNCTION,
		style: STYLE,
		styles: STYLES,
		pagesetup: PAGESETUP,
		validation: VALIDATION,
	},
	globals: {
		ThisWorkbook: WORKBOOK,
		ActiveWorkbook: WORKBOOK,
		ActiveSheet: WORKSHEET,
		ActiveCell: RANGE,
		Selection: RANGE,
		Application: APPLICATION,
		Cells: RANGE,
		Range: RANGE,
		Workbooks: WORKBOOKS,
		Worksheets: WORKSHEETS,
		Sheets: SHEETS,
	},
	constants: EXCEL_HOST_ENUM_CONSTANTS,
	types: {
		[APPLICATION]: {
			displayName: 'Application',
			members: mergeHostMembers([
				p('ActiveCell', RANGE),
				p('ActiveChart', CHART),
				p('ActiveSheet', WORKSHEET),
				p('ActiveWindow', WINDOW),
				p('ActiveWorkbook', WORKBOOK),
				p('AddIns'),
				p('Application', APPLICATION),
				p('Calculation'),
				p('Caption'),
				p('Cells', RANGE),
				p('Charts', CHARTS),
				p('Columns', RANGE),
				p('CommandBars'),
				p('Cursor'),
				p('DisplayAlerts'),
				p('DisplayFormulaBar'),
				p('DisplayFullScreen'),
				p('EnableEvents'),
				p('Height'),
				p('International'),
				p('Name'),
				p('Names', NAMES),
				p('Parent'),
				p('Path'),
				p('Range', RANGE),
				p('Rows', RANGE),
				p('ScreenUpdating'),
				p('Selection'),
				p('Sheets', SHEETS),
				p('StatusBar'),
				p('ThisWorkbook', WORKBOOK),
				p('UserName'),
				p('VBE'),
				p('Version'),
				p('Visible'),
				p('Width'),
				p('Windows', WINDOWS),
				p('Workbooks', WORKBOOKS),
				p('WorksheetFunction', WORKSHEETFUNCTION),
				p('Worksheets', WORKSHEETS),
				m('Calculate'),
				m('ConvertFormula'),
				m('Evaluate'),
				m('GetOpenFilename'),
				m('GetSaveAsFilename'),
				m('Goto'),
				m('InputBox'),
				m('Intersect', RANGE),
				m('OnTime'),
				m('Quit'),
				m('Run'),
				m('SendKeys'),
				m('Undo'),
				m('Union', RANGE),
				m('Volatile'),
				m('Wait'),
			], referenceMembers('Application')),
		},
		[WORKBOOK]: {
			displayName: 'Workbook',
			exhaustive: true,
			members: mergeHostMembers([
				p('ActiveChart', CHART),
				p('ActiveSheet', WORKSHEET),
				p('Application', APPLICATION),
				p('Charts', CHARTS),
				p('CodeName'),
				p('Colors'),
				p('CommandBars'),
				p('Connections'),
				p('Creator'),
				p('FullName'),
				p('HasPassword'),
				p('IsAddin'),
				p('Name'),
				p('Names', NAMES),
				p('Parent', APPLICATION),
				p('Path'),
				p('PivotTables', PIVOTTABLES),
				p('ReadOnly'),
				p('Saved'),
				p('Sheets', SHEETS),
				p('Styles', STYLES),
				p('VBProject'),
				p('Windows', WINDOWS),
				p('Worksheets', WORKSHEETS),
				m('Activate'),
				m('Close'),
				m('FollowHyperlink'),
				m('PrintOut'),
				m('PrintPreview'),
				m('Protect'),
				m('RefreshAll'),
				m('Save'),
				m('SaveAs'),
				m('SaveCopyAs'),
				m('SendMail'),
				m('Unprotect'),
			], EXCEL_WORKBOOK_REFERENCE_MEMBERS),
		},
		[WORKSHEET]: {
			displayName: 'Worksheet',
			exhaustive: true,
			members: mergeHostMembers([
				p('Application', APPLICATION),
				p('AutoFilter'),
				p('Cells', RANGE),
				p('CodeName'),
				p('Columns', RANGE),
				p('Comments', COMMENTS),
				p('Creator'),
				p('Hyperlinks', HYPERLINKS),
				p('Index'),
				p('ListObjects', LISTOBJECTS),
				p('Name'),
				p('Names', NAMES),
				p('Next', WORKSHEET),
				p('Outline'),
				p('PageSetup', PAGESETUP),
				p('Parent', WORKBOOK),
				p('Previous', WORKSHEET),
				p('Protection'),
				p('QueryTables'),
				p('Range', RANGE),
				p('Rows', RANGE),
				p('ScrollArea'),
				p('Shapes', SHAPES),
				p('Sort'),
				p('Tab'),
				p('Type'),
				p('UsedRange', RANGE),
				p('Visible'),
				m('Activate'),
				m('Buttons'),
				m('Calculate'),
				m('ChartObjects', CHARTOBJECTS),
				m('CheckSpelling'),
				m('Copy'),
				m('Delete'),
				m('Evaluate'),
				m('Move'),
				m('Paste'),
				m('PasteSpecial'),
				m('PivotTables', PIVOTTABLES),
				m('PrintOut'),
				m('PrintPreview'),
				m('Protect'),
				m('Select'),
				m('ShowAllData'),
				m('Unprotect'),
			], referenceMembers('Worksheet')),
		},
		[RANGE]: {
			displayName: 'Range',
			members: mergeHostMembers([
				p('Address'),
				p('Application', APPLICATION),
				p('Areas', AREAS),
				p('Borders', BORDERS),
				p('Cells', RANGE),
				p('Column'),
				p('Columns', RANGE),
				p('ColumnWidth'),
				p('Comment', COMMENT),
				p('Count'),
				p('CurrentRegion', RANGE),
				p('End', RANGE),
				p('EntireColumn', RANGE),
				p('EntireRow', RANGE),
				p('Font', FONT),
				p('Formula'),
				p('FormulaR1C1'),
				p('Height'),
				p('Hidden'),
				p('Hyperlinks', HYPERLINKS),
				p('HorizontalAlignment'),
				p('Interior', INTERIOR),
				p('Item'),
				p('Left'),
				p('ListObject', LISTOBJECT),
				p('Locked'),
				p('MergeArea', RANGE),
				p('MergeCells'),
				p('Name', NAME),
				p('Next', RANGE),
				p('NumberFormat'),
				p('Offset', RANGE),
				p('Orientation'),
				p('Parent', WORKSHEET),
				p('Previous', RANGE),
				p('Range', RANGE),
				p('Resize', RANGE),
				p('Row'),
				p('RowHeight'),
				p('Rows', RANGE),
				p('Style', STYLE),
				p('Text'),
				p('Top'),
				p('Validation', VALIDATION),
				p('Value'),
				p('Value2'),
				p('VerticalAlignment'),
				p('Width'),
				p('Worksheet', WORKSHEET),
				p('WrapText'),
				m('Activate'),
				m('AddComment'),
				m('AdvancedFilter'),
				m('AutoFill'),
				m('AutoFilter'),
				m('AutoFit'),
				m('BorderAround'),
				m('Calculate'),
				m('Clear'),
				m('ClearContents'),
				m('ClearFormats'),
				m('Copy'),
				m('Cut'),
				m('Delete'),
				m('Find', RANGE),
				m('FindNext', RANGE),
				m('Insert'),
				m('Merge'),
				m('PasteSpecial'),
				m('Replace'),
				m('Select'),
				m('Sort'),
				m('SpecialCells', RANGE),
				m('TextToColumns'),
				m('UnMerge'),
			], referenceMembers('Range')),
		},
		[WORKBOOKS]: {
			displayName: 'Workbooks',
			members: mergeHostMembers([
				p('Application', APPLICATION),
				p('Count'),
				p('Creator'),
				p('Item', WORKBOOK),
				p('Parent', APPLICATION),
				m('Add', WORKBOOK),
				m('Close'),
				m('Open', WORKBOOK),
				m('OpenDatabase', WORKBOOK),
				m('OpenText'),
				m('OpenXML', WORKBOOK),
			], referenceMembers('Workbooks')),
		},
		[WORKSHEETS]: {
			displayName: 'Worksheets',
			members: mergeHostMembers([
				p('Application', APPLICATION),
				p('Count'),
				p('Creator'),
				p('Item', WORKSHEET),
				p('Parent', WORKBOOK),
				p('Visible'),
				m('Add', WORKSHEET),
				m('Add2', WORKSHEET),
				m('Copy'),
				m('Delete'),
				m('FillAcrossSheets'),
				m('Move'),
				m('PrintOut'),
				m('PrintPreview'),
				m('Select'),
			], referenceMembers('Worksheets')),
		},
		[SHEETS]: {
			displayName: 'Sheets',
			members: mergeHostMembers([
				p('Application', APPLICATION),
				p('Count'),
				p('Creator'),
				// Sheets can contain worksheets or chart sheets; completion merges both.
				pAny('Item', [WORKSHEET, CHART]),
				p('Parent', WORKBOOK),
				p('Visible'),
				m('Add'),
				m('Copy'),
				m('Delete'),
				m('FillAcrossSheets'),
				m('Move'),
				m('PrintOut'),
				m('PrintPreview'),
				m('Select'),
			], referenceMembers('Sheets')),
		},
		[WINDOW]: {
			displayName: 'Window',
			members: [
				p('ActiveCell', RANGE),
				p('ActiveChart', CHART),
				p('ActivePane'),
				p('ActiveSheet', WORKSHEET),
				p('Application', APPLICATION),
				p('Caption'),
				p('DisplayGridlines'),
				p('DisplayHeadings'),
				p('FreezePanes'),
				p('Height'),
				p('Index'),
				p('Left'),
				p('Panes'),
				p('Parent', APPLICATION),
				p('RangeSelection', RANGE),
				p('ScrollColumn'),
				p('ScrollRow'),
				p('SelectedSheets'),
				p('Selection'),
				p('Split'),
				p('SplitColumn'),
				p('SplitRow'),
				p('TabRatio'),
				p('Top'),
				p('Type'),
				p('UsableHeight'),
				p('UsableWidth'),
				p('Visible'),
				p('VisibleRange', RANGE),
				p('Width'),
				p('WindowNumber'),
				p('WindowState'),
				p('Zoom'),
				m('Activate'),
				m('ActivateNext'),
				m('ActivatePrevious'),
				m('Close'),
				m('LargeScroll'),
				m('NewWindow', WINDOW),
				m('PrintOut'),
				m('PrintPreview'),
				m('ScrollIntoView'),
				m('SmallScroll'),
			],
		},
		[WINDOWS]: {
			displayName: 'Windows',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', WINDOW),
				p('Parent'),
				m('Arrange'),
				m('BreakSideBySide'),
				m('CompareSideBySideWith'),
			],
		},
		[NAME]: {
			displayName: 'Name',
			members: [
				p('Application', APPLICATION),
				p('Category'),
				p('Comment'),
				p('Index'),
				p('MacroType'),
				p('Name'),
				p('Parent'),
				p('RefersTo'),
				p('RefersToLocal'),
				p('RefersToR1C1'),
				p('RefersToRange', RANGE),
				p('ShortcutKey'),
				p('Value'),
				p('Visible'),
				m('Delete'),
			],
		},
		[NAMES]: {
			displayName: 'Names',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', NAME),
				m('Item', NAME),
			],
		},
		[COMMENT]: {
			displayName: 'Comment',
			members: [
				p('Application', APPLICATION),
				p('Author'),
				p('Parent'),
				p('Shape', SHAPE),
				p('Visible'),
				m('Delete'),
				m('Next', COMMENT),
				m('Previous', COMMENT),
				m('Text'),
			],
		},
		[COMMENTS]: {
			displayName: 'Comments',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', COMMENT),
			],
		},
		[LISTOBJECT]: {
			displayName: 'ListObject',
			members: [
				p('Active'),
				p('Application', APPLICATION),
				p('AutoFilter'),
				p('Comment'),
				p('DataBodyRange', RANGE),
				p('DisplayName'),
				p('HeaderRowRange', RANGE),
				p('InsertRowRange', RANGE),
				p('ListColumns', LISTCOLUMNS),
				p('ListRows', LISTROWS),
				p('Name'),
				p('Parent'),
				p('QueryTable'),
				p('Range', RANGE),
				p('ShowAutoFilter'),
				p('ShowHeaders'),
				p('ShowTotals'),
				p('Sort'),
				p('SourceType'),
				p('TableStyle'),
				p('TotalsRowRange', RANGE),
				m('Delete'),
				m('Publish'),
				m('Refresh'),
				m('Resize'),
				m('Unlink'),
				m('Unlist'),
			],
		},
		[LISTOBJECTS]: {
			displayName: 'ListObjects',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', LISTOBJECT),
				m('Item', LISTOBJECT),
			],
		},
		[LISTROW]: {
			displayName: 'ListRow',
			members: [
				p('Application', APPLICATION),
				p('Index'),
				p('Parent'),
				p('Range', RANGE),
				m('Delete'),
			],
		},
		[LISTROWS]: {
			displayName: 'ListRows',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', LISTROW),
				m('Item', LISTROW),
			],
		},
		[LISTCOLUMN]: {
			displayName: 'ListColumn',
			members: [
				p('Application', APPLICATION),
				p('DataBodyRange', RANGE),
				p('Index'),
				p('Name'),
				p('Parent'),
				p('Range', RANGE),
				p('Total', RANGE),
				p('TotalsCalculation'),
				m('Delete'),
			],
		},
		[LISTCOLUMNS]: {
			displayName: 'ListColumns',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', LISTCOLUMN),
				m('Item', LISTCOLUMN),
			],
		},
		[PIVOTTABLE]: {
			displayName: 'PivotTable',
			members: [
				p('Application', APPLICATION),
				p('ColumnGrand'),
				p('DataBodyRange', RANGE),
				p('Name'),
				p('PageRange', RANGE),
				p('Parent'),
				p('RefreshDate'),
				p('RowGrand'),
				p('RowRange', RANGE),
				p('SaveData'),
				p('SourceData'),
				p('TableRange1', RANGE),
				p('TableRange2', RANGE),
				p('Value'),
				m('AddDataField'),
				m('AddFields'),
				m('ClearTable'),
				m('GetData'),
				m('GetPivotData', RANGE),
				m('PivotFields'),
				m('RefreshTable'),
				m('Update'),
			],
		},
		[PIVOTTABLES]: {
			displayName: 'PivotTables',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', PIVOTTABLE),
				m('Item', PIVOTTABLE),
			],
		},
		[CHART]: {
			displayName: 'Chart',
			members: [
				p('Application', APPLICATION),
				p('ChartArea'),
				p('ChartTitle'),
				p('ChartType'),
				p('HasLegend'),
				p('HasTitle'),
				p('Legend'),
				p('Name'),
				p('PageSetup', PAGESETUP),
				p('Parent'),
				p('PlotArea'),
				p('Shapes', SHAPES),
				p('Type'),
				p('Visible'),
				m('Activate'),
				m('ApplyDataLabels'),
				m('Axes'),
				m('ChartObjects', CHARTOBJECTS),
				m('ChartWizard'),
				m('Copy'),
				m('Delete'),
				m('Export'),
				m('Location'),
				m('Paste'),
				m('Refresh'),
				m('Select'),
				m('SeriesCollection'),
				m('SetSourceData'),
			],
		},
		[CHARTS]: {
			displayName: 'Charts',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', CHART),
				p('Parent'),
				p('Visible'),
				m('Add2', CHART),
				m('Copy'),
				m('Delete'),
				m('Select'),
			],
		},
		[CHARTOBJECT]: {
			displayName: 'ChartObject',
			members: [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Chart', CHART),
				p('Height'),
				p('Index'),
				p('Left'),
				p('Locked'),
				p('Name'),
				p('Parent'),
				p('Placement'),
				p('Top'),
				p('TopLeftCell', RANGE),
				p('Visible'),
				p('Width'),
				m('Activate'),
				m('BringToFront'),
				m('Copy'),
				m('CopyPicture'),
				m('Cut'),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
				m('SendToBack'),
			],
		},
		[CHARTOBJECTS]: {
			displayName: 'ChartObjects',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Height'),
				p('Left'),
				p('Parent'),
				p('Top'),
				p('Visible'),
				p('Width'),
				m('Add', CHARTOBJECT),
				m('Delete'),
				m('Item', CHARTOBJECT),
				m('Select'),
			],
		},
		[SHAPE]: {
			displayName: 'Shape',
			members: [
				p('AlternativeText'),
				p('Application', APPLICATION),
				p('AutoShapeType'),
				p('BottomRightCell', RANGE),
				p('Chart', CHART),
				p('Fill'),
				p('Height'),
				p('Left'),
				p('Line'),
				p('LockAspectRatio'),
				p('Name'),
				p('Parent'),
				p('Placement'),
				p('Rotation'),
				p('TextFrame'),
				p('TextFrame2'),
				p('Title'),
				p('Top'),
				p('TopLeftCell', RANGE),
				p('Type'),
				p('Visible'),
				p('Width'),
				p('ZOrderPosition'),
				m('Apply'),
				m('Copy'),
				m('Cut'),
				m('Delete'),
				m('Duplicate'),
				m('Flip'),
				m('IncrementLeft'),
				m('IncrementRotation'),
				m('IncrementTop'),
				m('ScaleHeight'),
				m('ScaleWidth'),
				m('Select'),
				m('ZOrder'),
			],
		},
		[SHAPES]: {
			displayName: 'Shapes',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('Range'),
				m('AddChart2', SHAPE),
				m('AddConnector', SHAPE),
				m('AddLabel', SHAPE),
				m('AddLine', SHAPE),
				m('AddPicture', SHAPE),
				m('AddShape', SHAPE),
				m('AddTextbox', SHAPE),
				m('Item', SHAPE),
				m('SelectAll'),
			],
		},
		[FONT]: {
			displayName: 'Font',
			members: [
				p('Application', APPLICATION),
				p('Background'),
				p('Bold'),
				p('Color'),
				p('ColorIndex'),
				p('FontStyle'),
				p('Italic'),
				p('Name'),
				p('Parent'),
				p('Size'),
				p('Strikethrough'),
				p('Subscript'),
				p('Superscript'),
				p('ThemeColor'),
				p('ThemeFont'),
				p('TintAndShade'),
				p('Underline'),
			],
		},
		[INTERIOR]: {
			displayName: 'Interior',
			members: [
				p('Application', APPLICATION),
				p('Color'),
				p('ColorIndex'),
				p('Gradient'),
				p('InvertIfNegative'),
				p('Parent'),
				p('Pattern'),
				p('PatternColor'),
				p('PatternColorIndex'),
				p('ThemeColor'),
				p('TintAndShade'),
			],
		},
		[BORDER]: {
			displayName: 'Border',
			members: [
				p('Application', APPLICATION),
				p('Color'),
				p('ColorIndex'),
				p('LineStyle'),
				p('Parent'),
				p('ThemeColor'),
				p('TintAndShade'),
				p('Weight'),
			],
		},
		[BORDERS]: {
			displayName: 'Borders',
			members: [
				p('Application', APPLICATION),
				p('Color'),
				p('ColorIndex'),
				p('Count'),
				p('Item', BORDER),
				p('LineStyle'),
				p('Parent'),
				p('ThemeColor'),
				p('TintAndShade'),
				p('Value'),
				p('Weight'),
			],
		},
		[AREAS]: {
			displayName: 'Areas',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', RANGE),
				p('Parent'),
			],
		},
		[HYPERLINK]: {
			displayName: 'Hyperlink',
			members: [
				p('Address'),
				p('Application', APPLICATION),
				p('EmailSubject'),
				p('Name'),
				p('Parent'),
				p('Range', RANGE),
				p('ScreenTip'),
				p('Shape', SHAPE),
				p('SubAddress'),
				p('TextToDisplay'),
				p('Type'),
				m('AddToFavorites'),
				m('CreateNewDocument'),
				m('Delete'),
				m('Follow'),
			],
		},
		[HYPERLINKS]: {
			displayName: 'Hyperlinks',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', HYPERLINK),
				p('Parent'),
				m('Add', HYPERLINK),
				m('Delete'),
			],
		},
		[WORKSHEETFUNCTION]: {
			displayName: 'WorksheetFunction',
			members: [
				p('Application', APPLICATION),
				p('Parent'),
				m('Average'),
				m('AverageIf'),
				m('AverageIfs'),
				m('Ceiling'),
				m('Choose'),
				m('Clean'),
				m('Concat'),
				m('Count'),
				m('CountA'),
				m('CountBlank'),
				m('CountIf'),
				m('CountIfs'),
				m('Floor'),
				m('HLookup'),
				m('IfError'),
				m('Index'),
				m('Large'),
				m('Lookup'),
				m('Match'),
				m('Max'),
				m('MaxIfs'),
				m('Median'),
				m('Min'),
				m('MinIfs'),
				m('Mode'),
				m('Power'),
				m('Product'),
				m('Proper'),
				m('Rank'),
				m('Replace'),
				m('Round'),
				m('RoundDown'),
				m('RoundUp'),
				m('Search'),
				m('Small'),
				m('StDev'),
				m('Substitute'),
				m('Subtotal'),
				m('Sum'),
				m('SumIf'),
				m('SumIfs'),
				m('SumProduct'),
				m('TextJoin'),
				m('Transpose'),
				m('Trim'),
				m('VLookup'),
				m('Var'),
				m('XLookup'),
				m('XMatch'),
			],
		},
		[STYLE]: {
			displayName: 'Style',
			members: [
				p('Application', APPLICATION),
				p('Borders', BORDERS),
				p('BuiltIn'),
				p('Font', FONT),
				p('HorizontalAlignment'),
				p('IncludeAlignment'),
				p('IncludeBorder'),
				p('IncludeFont'),
				p('IncludeNumber'),
				p('IncludePatterns'),
				p('IncludeProtection'),
				p('Interior', INTERIOR),
				p('Locked'),
				p('MergeCells'),
				p('Name'),
				p('NumberFormat'),
				p('Parent'),
				p('Value'),
				p('VerticalAlignment'),
				p('WrapText'),
				m('Delete'),
			],
		},
		[STYLES]: {
			displayName: 'Styles',
			members: [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', STYLE),
				p('Parent'),
				m('Add', STYLE),
				m('Merge'),
			],
		},
		[PAGESETUP]: {
			displayName: 'PageSetup',
			members: [
				p('Application', APPLICATION),
				p('BlackAndWhite'),
				p('BottomMargin'),
				p('CenterFooter'),
				p('CenterHeader'),
				p('CenterHorizontally'),
				p('CenterVertically'),
				p('Draft'),
				p('FirstPageNumber'),
				p('FitToPagesTall'),
				p('FitToPagesWide'),
				p('FooterMargin'),
				p('HeaderMargin'),
				p('LeftFooter'),
				p('LeftHeader'),
				p('LeftMargin'),
				p('Order'),
				p('Orientation'),
				p('PaperSize'),
				p('Parent'),
				p('PrintArea'),
				p('PrintGridlines'),
				p('PrintHeadings'),
				p('RightFooter'),
				p('RightHeader'),
				p('RightMargin'),
				p('TopMargin'),
				p('Zoom'),
			],
		},
		[VALIDATION]: {
			displayName: 'Validation',
			members: [
				p('AlertStyle'),
				p('Application', APPLICATION),
				p('ErrorMessage'),
				p('ErrorTitle'),
				p('Formula1'),
				p('Formula2'),
				p('IgnoreBlank'),
				p('InCellDropdown'),
				p('InputMessage'),
				p('InputTitle'),
				p('Operator'),
				p('Parent'),
				p('ShowError'),
				p('ShowInput'),
				p('Type'),
				p('Value'),
				m('Add'),
				m('Delete'),
				m('Modify'),
			],
		},
	},
	// Verified call signatures transcribed from the Office VBA object-model
	// reference (learn.microsoft.com). Parameter lists are reproduced exactly;
	// where a method accepts a large variadic tail (e.g. Application.Run takes
	// Arg1..Arg30) only the leading, commonly-used parameters are listed rather
	// than inventing a synthetic "..." token.
	memberSignatures: {
		[WORKBOOKS]: {
			open:
				'Open(Filename As String, [UpdateLinks], [ReadOnly], [Format], ' +
				'[Password], [WriteResPassword], [IgnoreReadOnlyRecommended], ' +
				'[Origin], [Delimiter], [Editable], [Notify], [Converter], ' +
				'[AddToMru], [Local], [CorruptLoad]) As Workbook',
			add: 'Add([Template]) As Workbook',
			item: 'Item(Index) As Workbook',
		},
		[WORKSHEETS]: {
			add: 'Add([Before], [After], [Count], [Type]) As Worksheet',
			item: 'Item(Index) As Worksheet',
		},
		[SHEETS]: {
			add: 'Add([Before], [After], [Count], [Type])',
			item: 'Item(Index)',
		},
		[WORKBOOK]: {
			close: 'Close([SaveChanges], [Filename], [RouteWorkbook])',
			saveas:
				'SaveAs([Filename], [FileFormat], [Password], [WriteResPassword], ' +
				'[ReadOnlyRecommended], [CreateBackup], ' +
				'[AccessMode As XlSaveAsAccessMode = xlNoChange], ' +
				'[ConflictResolution], [AddToMru], [TextCodepage], ' +
				'[TextVisualLayout], [Local])',
			protect: 'Protect([Password], [Structure], [Windows])',
			unprotect: 'Unprotect([Password])',
			printout:
				'PrintOut([From], [To], [Copies], [Preview], [ActivePrinter], ' +
				'[PrintToFile], [Collate], [PrToFileName], [IgnorePrintAreas])',
		},
		[WORKSHEET]: {
			buttons: 'Buttons([Index])',
			range: 'Range(Cell1, [Cell2]) As Range',
			cells: 'Cells([RowIndex], [ColumnIndex]) As Range',
			protect:
				'Protect([Password], [DrawingObjects], [Contents], [Scenarios], ' +
				'[UserInterfaceOnly], [AllowFormattingCells], ' +
				'[AllowFormattingColumns], [AllowFormattingRows], ' +
				'[AllowInsertingColumns], [AllowInsertingRows], ' +
				'[AllowInsertingHyperlinks], [AllowDeletingColumns], ' +
				'[AllowDeletingRows], [AllowSorting], [AllowFiltering], ' +
				'[AllowUsingPivotTables])',
			unprotect: 'Unprotect([Password])',
			copy: 'Copy([Before], [After])',
			move: 'Move([Before], [After])',
		},
		[RANGE]: {
			range: 'Range(Cell1, [Cell2]) As Range',
			cells: 'Cells([RowIndex], [ColumnIndex]) As Range',
			offset: 'Offset([RowOffset], [ColumnOffset]) As Range',
			resize: 'Resize([RowSize], [ColumnSize]) As Range',
			find:
				'Find(What, [After], [LookIn], [LookAt], [SearchOrder], ' +
				'[SearchDirection], [MatchCase], [MatchByte], [SearchFormat]) As Range',
			specialcells: 'SpecialCells(Type As XlCellType, [Value]) As Range',
			autofilter:
				'AutoFilter([Field], [Criteria1], [Operator As XlAutoFilterOperator = xlAnd], ' +
				'[Criteria2], [VisibleDropDown])',
			pastespecial:
				'PasteSpecial([Paste As XlPasteType = xlPasteAll], ' +
				'[Operation As XlPasteSpecialOperation = xlPasteSpecialOperationNone], ' +
				'[SkipBlanks], [Transpose])',
		},
		[APPLICATION]: {
			inputbox:
				'InputBox(Prompt, [Title], [Default], [Left], [Top], [HelpFile], ' +
				'[HelpContextID], [Type]) As Variant',
			intersect:
				'Intersect(Arg1 As Range, Arg2 As Range, [Arg3], [Arg4]) As Range',
			union: 'Union(Arg1 As Range, Arg2 As Range, [Arg3], [Arg4]) As Range',
			ontime:
				'OnTime(EarliestTime, Procedure As String, [LatestTime], [Schedule])',
			goto: 'Goto([Reference], [Scroll])',
			wait: 'Wait(Time) As Boolean',
		},
		[NAMES]: {
			add:
				'Add([Name], [RefersTo], [Visible], [MacroType], [ShortcutKey], ' +
				'[Category], [NameLocal], [RefersToLocal], [CategoryLocal], ' +
				'[RefersToR1C1], [RefersToR1C1Local]) As Name',
			item: 'Item(Index) As Name',
		},
		[LISTOBJECTS]: {
			add:
				'Add([SourceType As XlListObjectSourceType], [Source], ' +
				'[LinkSource], [XlListObjectHasHeaders As XlYesNoGuess], ' +
				'[Destination], [TableStyleName]) As ListObject',
			item: 'Item(Index) As ListObject',
		},
		[HYPERLINKS]: {
			add:
				'Add(Anchor As Object, [Address], [SubAddress], [ScreenTip], ' +
				'[TextToDisplay]) As Hyperlink',
		},
		[CHARTOBJECTS]: {
			add: 'Add(Left As Double, Top As Double, Width As Double, Height As Double) As ChartObject',
			item: 'Item(Index) As ChartObject',
		},
	},
};
