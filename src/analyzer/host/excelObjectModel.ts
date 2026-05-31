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

export type HostMemberKind = 'property' | 'method' | 'event';

export interface HostMember {
	name: string;
	kind: HostMemberKind;
	/** Qualified host type this member returns, when stable (for chaining). */
	returns?: string;
}

export interface HostType {
	/** Bare display name, e.g. "Worksheet". */
	displayName: string;
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
}

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const RANGE = 'Excel.Range';
const APPLICATION = 'Excel.Application';

function p(name: string, returns?: string): HostMember {
	return { name, kind: 'property', returns };
}
function m(name: string, returns?: string): HostMember {
	return { name, kind: 'method', returns };
}

export const EXCEL_OBJECT_MODEL: HostObjectModel = {
	source: 'Office VBA object-model reference (learn.microsoft.com), verified 2026-05-30',
	aliases: {
		workbook: WORKBOOK,
		worksheet: WORKSHEET,
		range: RANGE,
		application: APPLICATION,
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
	},
	types: {
		[APPLICATION]: {
			displayName: 'Application',
			members: [
				p('ActiveCell', RANGE),
				p('ActiveChart'),
				p('ActiveSheet', WORKSHEET),
				p('ActiveWindow'),
				p('ActiveWorkbook', WORKBOOK),
				p('AddIns'),
				p('Application', APPLICATION),
				p('Calculation'),
				p('Caption'),
				p('Cells', RANGE),
				p('Charts'),
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
				p('Names'),
				p('Parent'),
				p('Path'),
				p('Range', RANGE),
				p('Rows', RANGE),
				p('ScreenUpdating'),
				p('Selection'),
				p('Sheets'),
				p('StatusBar'),
				p('ThisWorkbook', WORKBOOK),
				p('UserName'),
				p('VBE'),
				p('Version'),
				p('Visible'),
				p('Width'),
				p('Windows'),
				p('Workbooks'),
				p('WorksheetFunction'),
				p('Worksheets'),
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
			],
		},
		[WORKBOOK]: {
			displayName: 'Workbook',
			members: [
				p('ActiveChart'),
				p('ActiveSheet', WORKSHEET),
				p('Application', APPLICATION),
				p('Charts'),
				p('CodeName'),
				p('Colors'),
				p('CommandBars'),
				p('Connections'),
				p('Creator'),
				p('FullName'),
				p('HasPassword'),
				p('IsAddin'),
				p('Name'),
				p('Names'),
				p('Parent', APPLICATION),
				p('Path'),
				p('PivotTables'),
				p('ReadOnly'),
				p('Saved'),
				p('Sheets'),
				p('Styles'),
				p('VBProject'),
				p('Windows'),
				p('Worksheets'),
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
			],
		},
		[WORKSHEET]: {
			displayName: 'Worksheet',
			members: [
				p('Application', APPLICATION),
				p('AutoFilter'),
				p('Cells', RANGE),
				p('CodeName'),
				p('Columns', RANGE),
				p('Comments'),
				p('Creator'),
				p('Index'),
				p('ListObjects'),
				p('Name'),
				p('Names'),
				p('Next', WORKSHEET),
				p('Outline'),
				p('PageSetup'),
				p('Parent', WORKBOOK),
				p('Previous', WORKSHEET),
				p('Protection'),
				p('QueryTables'),
				p('Range', RANGE),
				p('Rows', RANGE),
				p('ScrollArea'),
				p('Shapes'),
				p('Sort'),
				p('Tab'),
				p('Type'),
				p('UsedRange', RANGE),
				p('Visible'),
				m('Activate'),
				m('Calculate'),
				m('ChartObjects'),
				m('CheckSpelling'),
				m('Copy'),
				m('Delete'),
				m('Evaluate'),
				m('Move'),
				m('Paste'),
				m('PasteSpecial'),
				m('PivotTables'),
				m('PrintOut'),
				m('PrintPreview'),
				m('Protect'),
				m('Select'),
				m('ShowAllData'),
				m('Unprotect'),
			],
		},
		[RANGE]: {
			displayName: 'Range',
			members: [
				p('Address'),
				p('Application', APPLICATION),
				p('Areas'),
				p('Borders'),
				p('Cells', RANGE),
				p('Column'),
				p('Columns', RANGE),
				p('ColumnWidth'),
				p('Comment'),
				p('Count'),
				p('CurrentRegion', RANGE),
				p('End', RANGE),
				p('EntireColumn', RANGE),
				p('EntireRow', RANGE),
				p('Font'),
				p('Formula'),
				p('FormulaR1C1'),
				p('Height'),
				p('Hidden'),
				p('HorizontalAlignment'),
				p('Interior'),
				p('Item'),
				p('Left'),
				p('ListObject'),
				p('Locked'),
				p('MergeArea', RANGE),
				p('MergeCells'),
				p('Name'),
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
				p('Style'),
				p('Text'),
				p('Top'),
				p('Validation'),
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
			],
		},
	},
};
