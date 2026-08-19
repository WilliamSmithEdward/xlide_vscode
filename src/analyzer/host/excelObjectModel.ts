// Excel host object-model metadata for member completion.
//
// SOURCE OF TRUTH: the member names below are transcribed from the official
// Microsoft Office VBA object-model reference (learn.microsoft.com/office/vba/
// api/excel.*), verified on 2026-05-30, plus generated Excel reference dumps
// checked into src/analyzer/host/excelReferenceMembers.ts. Promoted generated
// types feed completion, hover, signature help, and receiver chains; only the
// hard-diagnostic subset is marked exhaustive for member-not-found. The
// remaining curated types are common member-completion and receiver-chain
// helpers. Return types are provided only for members whose result type is stable and
// documented, to enable member-access chaining (e.g. ws.Range(...).Offset(...).).
//
// This data is host metadata, NOT VBA language grammar. Per
// docs/xlide_vba_language_service_roadmap.md it must never override core
// MS-VBAL language resolution. LLM-generated member lists are never used here.

import {
	EXCEL_REFERENCE_ENUM_CONSTANTS,
	EXCEL_REFERENCE_HARD_DIAGNOSTIC_TYPES,
	EXCEL_REFERENCE_MEMBER_SETS,
	EXCEL_REFERENCE_PROMOTED_TYPES,
	EXCEL_REFERENCE_PROVENANCE,
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
	/** Reference/source provenance for generated or promoted host metadata. */
	provenance?: string;
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
	/**
	 * Application display name for origin labels ("Excel host method",
	 * "Word type"). Absent means Excel: the default model when no host is
	 * named is Excel's (issue #28).
	 */
	hostName?: string;
	/** Qualified type name (e.g. "Excel.Range") -> type metadata. */
	types: Record<string, HostType>;
	/** Lowercased type name as written in `As <type>` -> qualified type. */
	aliases: Record<string, string>;
	/** Host-injected global identifier (canonical casing) -> qualified type. */
	globals: Record<string, string>;
	/**
	 * Qualified name (in `types`) of the host's hidden Global interface, whose
	 * members VBA calls bare - Word's InchesToPoints, Excel's Union (issue
	 * #34). Absent when the host has none modelled.
	 */
	globalType?: string;
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
const PIVOTFIELD = 'Excel.PivotField';
const PIVOTFIELDS = 'Excel.PivotFields';
const PIVOTITEM = 'Excel.PivotItem';
const PIVOTITEMS = 'Excel.PivotItems';
const PIVOTCACHE = 'Excel.PivotCache';
const PIVOTCACHES = 'Excel.PivotCaches';
const PIVOTFILTER = 'Excel.PivotFilter';
const PIVOTFILTERS = 'Excel.PivotFilters';
const CALCULATEDFIELDS = 'Excel.CalculatedFields';
const CALCULATEDITEMS = 'Excel.CalculatedItems';
const CUBEFIELD = 'Excel.CubeField';
const CUBEFIELDS = 'Excel.CubeFields';
const QUERYTABLE = 'Excel.QueryTable';
const QUERYTABLES = 'Excel.QueryTables';
const PARAMETER = 'Excel.Parameter';
const PARAMETERS = 'Excel.Parameters';
const CHART = 'Excel.Chart';
const CHARTS = 'Excel.Charts';
const CHARTOBJECT = 'Excel.ChartObject';
const CHARTOBJECTS = 'Excel.ChartObjects';
const SERIESCOLLECTION = 'Excel.SeriesCollection';
const FULLSERIESCOLLECTION = 'Excel.FullSeriesCollection';
const SERIES = 'Excel.Series';
const AXES = 'Excel.Axes';
const AXIS = 'Excel.Axis';
const AXISTITLE = 'Excel.AxisTitle';
const CHARTFORMAT = 'Excel.ChartFormat';
const TICKLABELS = 'Excel.TickLabels';
const GRIDLINES = 'Excel.Gridlines';
const DISPLAYUNITLABEL = 'Excel.DisplayUnitLabel';
const LEADERLINES = 'Excel.LeaderLines';
const ERRORBARS = 'Excel.ErrorBars';
const POINTS = 'Excel.Points';
const POINT = 'Excel.Point';
const TRENDLINES = 'Excel.Trendlines';
const TRENDLINE = 'Excel.Trendline';
const DATALABELS = 'Excel.DataLabels';
const DATALABEL = 'Excel.DataLabel';
const SHAPE = 'Excel.Shape';
const SHAPES = 'Excel.Shapes';
const SHAPERANGE = 'Excel.ShapeRange';
const GROUPSHAPES = 'Excel.GroupShapes';
const COMMENTTHREADED = 'Excel.CommentThreaded';
const COMMENTSTHREADED = 'Excel.CommentsThreaded';
const SORT = 'Excel.Sort';
const SORTFIELDS = 'Excel.SortFields';
const SORTFIELD = 'Excel.SortField';
const AUTOFILTER = 'Excel.AutoFilter';
const FILTERS = 'Excel.Filters';
const FILTER = 'Excel.Filter';
const ICON = 'Excel.Icon';
const ICONSET = 'Excel.IconSet';
const ICONSETS = 'Excel.IconSets';
const OLEOBJECT = 'Excel.OLEObject';
const OLEOBJECTS = 'Excel.OLEObjects';
const BUTTON = 'Excel.Button';
const BUTTONS = 'Excel.Buttons';
const CHECKBOX = 'Excel.CheckBox';
const CHECKBOXES = 'Excel.CheckBoxes';
const DROPDOWN = 'Excel.DropDown';
const DROPDOWNS = 'Excel.DropDowns';
const OPTIONBUTTON = 'Excel.OptionButton';
const OPTIONBUTTONS = 'Excel.OptionButtons';
const GROUPOBJECT = 'Excel.GroupObject';
const GROUPOBJECTS = 'Excel.GroupObjects';
const GROUPBOX = 'Excel.GroupBox';
const GROUPBOXES = 'Excel.GroupBoxes';
const LABEL = 'Excel.Label';
const LABELS = 'Excel.Labels';
const LISTBOX = 'Excel.ListBox';
const LISTBOXES = 'Excel.ListBoxes';
const SCROLLBAR = 'Excel.ScrollBar';
const SCROLLBARS = 'Excel.ScrollBars';
const SPINNER = 'Excel.Spinner';
const SPINNERS = 'Excel.Spinners';
const EDITBOX = 'Excel.EditBox';
const EDITBOXES = 'Excel.EditBoxes';
const TEXTBOX = 'Excel.TextBox';
const TEXTBOXES = 'Excel.TextBoxes';
const CONNECTIONS = 'Excel.Connections';
const WORKBOOKCONNECTION = 'Excel.WorkbookConnection';
const OLEDBCONNECTION = 'Excel.OLEDBConnection';
const ODBCCONNECTION = 'Excel.ODBCConnection';
const RANGES = 'Excel.Ranges';
const MODELCONNECTION = 'Excel.ModelConnection';
const WORKSHEETDATACONNECTION = 'Excel.WorksheetDataConnection';
const TEXTCONNECTION = 'Excel.TextConnection';
const DATAFEEDCONNECTION = 'Excel.DataFeedConnection';
const MODELTABLES = 'Excel.ModelTables';
const MODELTABLE = 'Excel.ModelTable';
const CALCULATEDMEMBER = 'Excel.CalculatedMember';
const CALCULATEDMEMBERS = 'Excel.CalculatedMembers';
const MODELTABLECOLUMN = 'Excel.ModelTableColumn';
const MODELTABLECOLUMNS = 'Excel.ModelTableColumns';
const SLICERCACHE = 'Excel.SlicerCache';
const SLICERCACHES = 'Excel.SlicerCaches';
const SLICER = 'Excel.Slicer';
const SLICERS = 'Excel.Slicers';
const SLICERITEM = 'Excel.SlicerItem';
const SLICERITEMS = 'Excel.SlicerItems';
const SLICERCACHELEVEL = 'Excel.SlicerCacheLevel';
const SLICERCACHELEVELS = 'Excel.SlicerCacheLevels';
const SLICERPIVOTTABLES = 'Excel.SlicerPivotTables';
const TIMELINESTATE = 'Excel.TimelineState';
const TIMELINEVIEWSTATE = 'Excel.TimelineViewState';
const FILLFORMAT = 'Excel.FillFormat';
const LINEFORMAT = 'Excel.LineFormat';
const TEXTFRAME = 'Excel.TextFrame';
const TEXTFRAME2 = 'Excel.TextFrame2';
const PICTUREFORMAT = 'Excel.PictureFormat';
const SHADOWFORMAT = 'Excel.ShadowFormat';
const THREEDFORMAT = 'Excel.ThreeDFormat';
const CONNECTORFORMAT = 'Excel.ConnectorFormat';
const CALLOUTFORMAT = 'Excel.CalloutFormat';
const SHAPENODES = 'Excel.ShapeNodes';
const SHAPENODE = 'Excel.ShapeNode';
const COLORFORMAT = 'Excel.ColorFormat';
const CHARTTITLE = 'Excel.ChartTitle';
const CHARTAREA = 'Excel.ChartArea';
const PLOTAREA = 'Excel.PlotArea';
const LEGEND = 'Excel.Legend';
const LEGENDENTRY = 'Excel.LegendEntry';
const LEGENDENTRIES = 'Excel.LegendEntries';
const LEGENDKEY = 'Excel.LegendKey';
const CHARTGROUP = 'Excel.ChartGroup';
const CHARTGROUPS = 'Excel.ChartGroups';
const DATATABLE = 'Excel.DataTable';
const WALLS = 'Excel.Walls';
const FLOOR = 'Excel.Floor';
const SERIESLINES = 'Excel.SeriesLines';
const DOWNBARS = 'Excel.DownBars';
const DROPLINES = 'Excel.DropLines';
const HILOLINES = 'Excel.HiLoLines';
const UPBARS = 'Excel.UpBars';
const DATABAR = 'Excel.Databar';
const COLORSCALE = 'Excel.ColorScale';
const COLORSCALECRITERIA = 'Excel.ColorScaleCriteria';
const COLORSCALECRITERION = 'Excel.ColorScaleCriterion';
const ICONSETCONDITION = 'Excel.IconSetCondition';
const ABOVEAVERAGE = 'Excel.AboveAverage';
const TOP10 = 'Excel.Top10';
const UNIQUEVALUES = 'Excel.UniqueValues';
const CELLFORMAT = 'Excel.CellFormat';
const DISPLAYFORMAT = 'Excel.DisplayFormat';
const FORMATCOLOR = 'Excel.FormatColor';
const CONDITIONVALUE = 'Excel.ConditionValue';
const DATABARBORDER = 'Excel.DataBarBorder';
const NEGATIVEBARFORMAT = 'Excel.NegativeBarFormat';
const ICONCRITERIA = 'Excel.IconCriteria';
const ICONCRITERION = 'Excel.IconCriterion';
const DRAWINGOBJECTS = 'Excel.DrawingObjects';
const DRAWING = 'Excel.Drawing';
const DRAWINGS = 'Excel.Drawings';
const PICTURE = 'Excel.Picture';
const PICTURES = 'Excel.Pictures';
const LINE = 'Excel.Line';
const LINES = 'Excel.Lines';
const RECTANGLE = 'Excel.Rectangle';
const RECTANGLES = 'Excel.Rectangles';
const OVAL = 'Excel.Oval';
const OVALS = 'Excel.Ovals';
const ARC = 'Excel.Arc';
const ARCS = 'Excel.Arcs';
const SPARKLINEGROUPS = 'Excel.SparklineGroups';
const SPARKLINEGROUP = 'Excel.SparklineGroup';
const SPARKLINE = 'Excel.Sparkline';
const SPARKPOINTS = 'Excel.SparkPoints';
const SPARKCOLOR = 'Excel.SparkColor';
const SPARKAXES = 'Excel.SparkAxes';
const SPARKHORIZONTALAXIS = 'Excel.SparkHorizontalAxis';
const SPARKVERTICALAXIS = 'Excel.SparkVerticalAxis';
const XMLMAP = 'Excel.XmlMap';
const XMLMAPS = 'Excel.XmlMaps';
const XPATH = 'Excel.XPath';
const PUBLISHOBJECT = 'Excel.PublishObject';
const PUBLISHOBJECTS = 'Excel.PublishObjects';
const WEBOPTIONS = 'Excel.WebOptions';
const DEFAULTWEBOPTIONS = 'Excel.DefaultWebOptions';
const XMLNAMESPACE = 'Excel.XmlNamespace';
const XMLNAMESPACES = 'Excel.XmlNamespaces';
const XMLSCHEMA = 'Excel.XmlSchema';
const XMLSCHEMAS = 'Excel.XmlSchemas';
const XMLDATABINDING = 'Excel.XmlDataBinding';
const FONT = 'Excel.Font';
const INTERIOR = 'Excel.Interior';
const BORDER = 'Excel.Border';
const BORDERS = 'Excel.Borders';
const AREAS = 'Excel.Areas';
const HYPERLINK = 'Excel.Hyperlink';
const HYPERLINKS = 'Excel.Hyperlinks';
const FORMATCONDITION = 'Excel.FormatCondition';
const FORMATCONDITIONS = 'Excel.FormatConditions';
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

/**
 * Merges constant tables, later sets winning name collisions
 * (case-insensitive). Exported so every generated host model can lay its own
 * library's constants over the shared Office table the way Excel does.
 */
export function mergeHostConstants(
	...sets: Array<Readonly<Record<string, HostConstant>>>
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

function promotedExcelReferenceProvenance(displayName: string): string {
	return EXCEL_REFERENCE_PROVENANCE[displayName] ?? `reference/excel/json/${displayName}.json`;
}

function promotedExcelReferenceExhaustive(displayName: string): boolean | undefined {
	return (EXCEL_REFERENCE_HARD_DIAGNOSTIC_TYPES as readonly string[]).includes(displayName)
		? true
		: undefined;
}

function promotedExcelHostType(
	displayName: string,
	primary: readonly HostMember[],
): HostType {
	return {
		displayName,
		provenance: promotedExcelReferenceProvenance(displayName),
		exhaustive: promotedExcelReferenceExhaustive(displayName),
		members: mergeHostMembers(primary, referenceMembers(displayName)),
	};
}

function promotedExcelReferenceAliases(): Record<string, string> {
	const aliases: Record<string, string> = {};
	for (const displayName of EXCEL_REFERENCE_PROMOTED_TYPES) {
		aliases[displayName.toLowerCase()] = `Excel.${displayName}`;
	}
	return aliases;
}

function promotedExcelHostTypes(displayNames: readonly string[]): Record<string, HostType> {
	const types: Record<string, HostType> = {};
	for (const displayName of displayNames) {
		types[`Excel.${displayName}`] = promotedExcelHostType(displayName, []);
	}
	return types;
}

let EXCEL_OBJECT_MODEL_CACHE: HostObjectModel | undefined;

/**
 * Lazily builds the Excel host object model on first access so the
 * generated-table merge (229 types, ~5,900 member copies) is not paid at
 * extension startup.
 */
export function getExcelObjectModel(): HostObjectModel {
	EXCEL_OBJECT_MODEL_CACHE ??= buildExcelObjectModel();
	return EXCEL_OBJECT_MODEL_CACHE;
}

const buildExcelObjectModel = (): HostObjectModel => ({
	source: `Office VBA object-model reference (learn.microsoft.com) + Excel COM type library, verified 2026-05-30; promoted Excel and Office reference enum constants; promoted Excel reference metadata for ${EXCEL_REFERENCE_PROMOTED_TYPES.join(', ')}; hard member-not-found diagnostics limited to ${EXCEL_REFERENCE_HARD_DIAGNOSTIC_TYPES.join(', ')}; Workbook dump ${EXCEL_WORKBOOK_REFERENCE_PROVENANCE}`,
	hostName: 'Excel',
	globalType: 'Excel.Global',
	aliases: {
		...promotedExcelReferenceAliases(),
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
		formatcondition: FORMATCONDITION,
		formatconditions: FORMATCONDITIONS,
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
	constants: mergeHostConstants(
		OFFICE_REFERENCE_ENUM_CONSTANTS,
		EXCEL_REFERENCE_ENUM_CONSTANTS,
	),
	types: {
		...promotedExcelHostTypes(EXCEL_REFERENCE_PROMOTED_TYPES),
		[APPLICATION]: {
			displayName: 'Application',
			provenance: promotedExcelReferenceProvenance('Application'),
			exhaustive: promotedExcelReferenceExhaustive('Application'),
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
				p('DefaultWebOptions', DEFAULTWEBOPTIONS),
				p('EnableEvents'),
				p('FindFormat', CELLFORMAT),
				p('Height'),
				p('International'),
				p('IconSets', ICONSETS),
				p('Name'),
				p('Names', NAMES),
				p('Parent'),
				p('Path'),
				p('Range', RANGE),
				p('ReplaceFormat', CELLFORMAT),
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
			provenance: promotedExcelReferenceProvenance('Workbook'),
			exhaustive: promotedExcelReferenceExhaustive('Workbook'),
			members: mergeHostMembers([
				p('ActiveChart', CHART),
				p('ActiveSheet', WORKSHEET),
				p('Application', APPLICATION),
				p('Charts', CHARTS),
				p('CodeName'),
				p('Colors'),
				p('CommandBars'),
				p('Connections', CONNECTIONS),
				p('Creator'),
				p('FullName'),
				p('HasPassword'),
				p('IsAddin'),
				p('Name'),
				p('Names', NAMES),
				p('Parent', APPLICATION),
				p('Path'),
				p('PivotCaches', PIVOTCACHES),
				p('PivotTables', PIVOTTABLES),
				p('PublishObjects', PUBLISHOBJECTS),
				p('ReadOnly'),
				p('Saved'),
				p('Sheets', SHEETS),
				p('SlicerCaches', SLICERCACHES),
				p('Styles', STYLES),
				p('VBProject'),
				p('WebOptions', WEBOPTIONS),
				p('Windows', WINDOWS),
				p('Worksheets', WORKSHEETS),
				p('XmlMaps', XMLMAPS),
				p('XmlNamespaces', XMLNAMESPACES),
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
			provenance: promotedExcelReferenceProvenance('Worksheet'),
			exhaustive: promotedExcelReferenceExhaustive('Worksheet'),
			members: mergeHostMembers([
				p('Application', APPLICATION),
				p('AutoFilter', AUTOFILTER),
				p('Cells', RANGE),
				p('CodeName'),
				p('Columns', RANGE),
				p('Comments', COMMENTS),
				p('CommentsThreaded', COMMENTSTHREADED),
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
				p('QueryTables', QUERYTABLES),
				p('Range', RANGE),
				p('Rows', RANGE),
				p('ScrollArea'),
				p('Shapes', SHAPES),
				p('Sort', SORT),
				p('Tab'),
				p('Type'),
				p('UsedRange', RANGE),
				p('Visible'),
				m('Arcs', ARCS),
				m('Activate'),
				m('Buttons', BUTTONS),
				m('Calculate'),
				m('ChartObjects', CHARTOBJECTS),
				m('CheckBoxes', CHECKBOXES),
				m('CheckSpelling'),
				m('Copy'),
				m('Delete'),
				m('DrawingObjects', DRAWINGOBJECTS),
				m('Drawings', DRAWINGS),
				m('DropDowns', DROPDOWNS),
				m('EditBoxes', EDITBOXES),
				m('Evaluate'),
				m('GroupBoxes', GROUPBOXES),
				m('Labels', LABELS),
				m('Lines', LINES),
				m('ListBoxes', LISTBOXES),
				m('Move'),
				m('OLEObjects', OLEOBJECTS),
				m('OptionButtons', OPTIONBUTTONS),
				m('Ovals', OVALS),
				m('Paste'),
				m('PasteSpecial'),
				m('Pictures', PICTURES),
				p('PivotTables', PIVOTTABLES),
				m('PrintOut'),
				m('PrintPreview'),
				m('Protect'),
				m('Rectangles', RECTANGLES),
				m('Select'),
				m('ShowAllData'),
				m('ScrollBars', SCROLLBARS),
				m('Spinners', SPINNERS),
				m('TextBoxes', TEXTBOXES),
				m('Unprotect'),
			], referenceMembers('Worksheet')),
		},
		[RANGE]: {
			displayName: 'Range',
			provenance: promotedExcelReferenceProvenance('Range'),
			exhaustive: promotedExcelReferenceExhaustive('Range'),
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
				p('CommentThreaded', COMMENTTHREADED),
				p('Count'),
				p('CurrentRegion', RANGE),
				p('DisplayFormat', DISPLAYFORMAT),
				p('End', RANGE),
				p('EntireColumn', RANGE),
				p('EntireRow', RANGE),
				p('Font', FONT),
				p('FormatConditions', FORMATCONDITIONS),
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
				p('SparklineGroups', SPARKLINEGROUPS),
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
				p('XPath', XPATH),
				m('Activate'),
				m('AddComment', COMMENT),
				m('AddCommentThreaded', COMMENTTHREADED),
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
			provenance: promotedExcelReferenceProvenance('Workbooks'),
			exhaustive: promotedExcelReferenceExhaustive('Workbooks'),
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
			provenance: promotedExcelReferenceProvenance('Worksheets'),
			exhaustive: promotedExcelReferenceExhaustive('Worksheets'),
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
			provenance: promotedExcelReferenceProvenance('Sheets'),
			exhaustive: promotedExcelReferenceExhaustive('Sheets'),
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
		[WINDOW]: promotedExcelHostType('Window', [
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
			]),
		[WINDOWS]: promotedExcelHostType('Windows', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', WINDOW),
				p('Parent'),
				m('Arrange'),
				m('BreakSideBySide'),
				m('CompareSideBySideWith'),
			]),
		[NAME]: promotedExcelHostType('Name', [
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
			]),
		[NAMES]: promotedExcelHostType('Names', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', NAME),
				m('Item', NAME),
			]),
		[COMMENT]: promotedExcelHostType('Comment', [
				p('Application', APPLICATION),
				p('Author'),
				p('Parent'),
				p('Shape', SHAPE),
				p('Visible'),
				m('Delete'),
				m('Next', COMMENT),
				m('Previous', COMMENT),
				m('Text'),
			]),
		[COMMENTS]: promotedExcelHostType('Comments', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', COMMENT),
			]),
		[COMMENTTHREADED]: promotedExcelHostType('CommentThreaded', [
				p('Application', APPLICATION),
				p('Parent'),
				p('Replies', COMMENTSTHREADED),
				m('AddReply', COMMENTTHREADED),
				m('Delete'),
				m('Next', COMMENTTHREADED),
				m('Previous', COMMENTTHREADED),
				m('Text'),
			]),
		[COMMENTSTHREADED]: promotedExcelHostType('CommentsThreaded', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', COMMENTTHREADED),
			]),
		[LISTOBJECT]: promotedExcelHostType('ListObject', [
				p('Active'),
				p('Application', APPLICATION),
				p('AutoFilter', AUTOFILTER),
				p('Comment'),
				p('DataBodyRange', RANGE),
				p('DisplayName'),
				p('HeaderRowRange', RANGE),
				p('InsertRowRange', RANGE),
				p('ListColumns', LISTCOLUMNS),
				p('ListRows', LISTROWS),
				p('Name'),
				p('Parent'),
				p('QueryTable', QUERYTABLE),
				p('Range', RANGE),
				p('ShowAutoFilter'),
				p('ShowHeaders'),
				p('ShowTotals'),
				p('Slicers', SLICERS),
				p('Sort', SORT),
				p('SourceType'),
				p('TableStyle'),
				p('TotalsRowRange', RANGE),
				p('XmlMap', XMLMAP),
				m('Delete'),
				m('Publish'),
				m('Refresh'),
				m('Resize'),
				m('Unlink'),
				m('Unlist'),
			]),
		[LISTOBJECTS]: promotedExcelHostType('ListObjects', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', LISTOBJECT),
				m('Item', LISTOBJECT),
			]),
		[LISTROW]: promotedExcelHostType('ListRow', [
				p('Application', APPLICATION),
				p('Index'),
				p('Parent'),
				p('Range', RANGE),
				m('Delete'),
			]),
		[LISTROWS]: promotedExcelHostType('ListRows', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', LISTROW),
				m('Item', LISTROW),
			]),
		[LISTCOLUMN]: promotedExcelHostType('ListColumn', [
				p('Application', APPLICATION),
				p('DataBodyRange', RANGE),
				p('Index'),
				p('Name'),
				p('Parent'),
				p('Range', RANGE),
				p('Total', RANGE),
				p('TotalsCalculation'),
				p('XPath', XPATH),
				m('Delete'),
			]),
		[LISTCOLUMNS]: promotedExcelHostType('ListColumns', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', LISTCOLUMN),
				m('Item', LISTCOLUMN),
			]),
		[PIVOTTABLE]: promotedExcelHostType('PivotTable', [
				p('Application', APPLICATION),
				p('CalculatedFields', CALCULATEDFIELDS),
				p('CalculatedItems', CALCULATEDITEMS),
				p('ColumnFields', PIVOTFIELDS),
				p('ColumnGrand'),
				p('CubeFields', CUBEFIELDS),
				p('DataBodyRange', RANGE),
				p('DataFields', PIVOTFIELDS),
				p('HiddenFields', PIVOTFIELDS),
				p('Name'),
				p('PageFields', PIVOTFIELDS),
				p('PageRange', RANGE),
				p('Parent'),
				p('PivotCache', PIVOTCACHE),
				p('PivotFields', PIVOTFIELDS),
				p('PivotFilters', PIVOTFILTERS),
				p('RefreshDate'),
				p('RowFields', PIVOTFIELDS),
				p('RowGrand'),
				p('RowRange', RANGE),
				p('SaveData'),
				p('SourceData'),
				p('TableRange1', RANGE),
				p('TableRange2', RANGE),
				p('Value'),
				p('VisibleFields', PIVOTFIELDS),
				m('AddDataField'),
				m('AddFields'),
				m('ClearTable'),
				m('GetData'),
				m('GetPivotData', RANGE),
				m('RefreshTable'),
				m('Update'),
			]),
		[PIVOTTABLES]: promotedExcelHostType('PivotTables', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', PIVOTTABLE),
				m('Item', PIVOTTABLE),
			]),
		[PIVOTFIELD]: promotedExcelHostType('PivotField', [
				p('Application', APPLICATION),
				p('ChildItems', PIVOTITEMS),
				p('HiddenItems', PIVOTITEMS),
				p('ParentItems', PIVOTITEMS),
				p('PivotFilters', PIVOTFILTERS),
				p('PivotItems', PIVOTITEMS),
				p('SourceRange', RANGE),
				p('VisibleItems', PIVOTITEMS),
				m('ClearAllFilters'),
			]),
		[PIVOTFIELDS]: promotedExcelHostType('PivotFields', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', PIVOTFIELD),
			]),
		[PIVOTITEM]: promotedExcelHostType('PivotItem', [
				p('Application', APPLICATION),
				p('ChildItems', PIVOTITEMS),
				p('DataRange', RANGE),
				p('LabelRange', RANGE),
				p('ParentItem', PIVOTITEM),
				p('ParentShowDetail', PIVOTITEM),
				p('Position'),
				p('SourceName'),
				p('Value'),
			]),
		[PIVOTITEMS]: promotedExcelHostType('PivotItems', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', PIVOTITEM),
			]),
		[PIVOTCACHE]: promotedExcelHostType('PivotCache', [
				p('Application', APPLICATION),
				p('Creator'),
				p('Index'),
				p('MemoryUsed'),
				p('Parent'),
				p('RecordCount'),
				m('CreatePivotTable', PIVOTTABLE),
				m('Refresh'),
			]),
		[PIVOTCACHES]: promotedExcelHostType('PivotCaches', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Create', PIVOTCACHE),
				m('Item', PIVOTCACHE),
			]),
		[PIVOTFILTER]: promotedExcelHostType('PivotFilter', [
				p('Application', APPLICATION),
				p('DataField', PIVOTFIELD),
				p('FilteringPivotField', PIVOTFIELD),
				p('MemberPropertyField', PIVOTFIELD),
				p('Parent'),
				p('Value1'),
				p('Value2'),
				m('Delete'),
			]),
		[PIVOTFILTERS]: promotedExcelHostType('PivotFilters', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', PIVOTFILTER),
				m('Item', PIVOTFILTER),
			]),
		[CALCULATEDFIELDS]: promotedExcelHostType('CalculatedFields', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', PIVOTFIELD),
				m('Item', PIVOTFIELD),
			]),
		[CALCULATEDITEMS]: promotedExcelHostType('CalculatedItems', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', PIVOTITEM),
				m('Item', PIVOTITEM),
			]),
		[CUBEFIELD]: promotedExcelHostType('CubeField', [
				p('Application', APPLICATION),
				p('Caption'),
				p('Name'),
				p('Parent'),
				m('CreatePivotFields'),
			]),
		[CUBEFIELDS]: promotedExcelHostType('CubeFields', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', CUBEFIELD),
			]),
		[QUERYTABLE]: promotedExcelHostType('QueryTable', [
				p('Application', APPLICATION),
				p('Destination', RANGE),
				p('ListObject', LISTOBJECT),
				p('Parameters', PARAMETERS),
				p('ResultRange', RANGE),
				p('Sort', SORT),
				p('WorkbookConnection', WORKBOOKCONNECTION),
				m('Refresh'),
			]),
		[QUERYTABLES]: promotedExcelHostType('QueryTables', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', QUERYTABLE),
				m('Add', QUERYTABLE),
				m('Item', QUERYTABLE),
			]),
		[PARAMETER]: promotedExcelHostType('Parameter', [
				p('Application', APPLICATION),
				p('Parent'),
				p('SourceRange', RANGE),
				m('SetParam'),
			]),
		[PARAMETERS]: promotedExcelHostType('Parameters', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', PARAMETER),
				m('Add', PARAMETER),
				m('Item', PARAMETER),
			]),
		[CONNECTIONS]: promotedExcelHostType('Connections', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', WORKBOOKCONNECTION),
				m('Add2', WORKBOOKCONNECTION),
				m('AddFromFile', WORKBOOKCONNECTION),
				m('Item', WORKBOOKCONNECTION),
			]),
		[WORKBOOKCONNECTION]: promotedExcelHostType('WorkbookConnection', [
				p('Application', APPLICATION),
				p('DataFeedConnection', DATAFEEDCONNECTION),
				p('ModelConnection', MODELCONNECTION),
				p('ModelTables', MODELTABLES),
				p('ODBCConnection', ODBCCONNECTION),
				p('OLEDBConnection', OLEDBCONNECTION),
				p('Ranges', RANGES),
				p('TextConnection', TEXTCONNECTION),
				p('WorksheetDataConnection', WORKSHEETDATACONNECTION),
				m('Delete'),
				m('Refresh'),
			]),
		[OLEDBCONNECTION]: promotedExcelHostType('OLEDBConnection', [
				p('Application', APPLICATION),
				p('CalculatedMembers', CALCULATEDMEMBERS),
				p('Parent'),
				m('CancelRefresh'),
				m('MakeConnection'),
				m('Reconnect'),
				m('Refresh'),
			]),
		[ODBCCONNECTION]: promotedExcelHostType('ODBCConnection', [
				p('Application', APPLICATION),
				p('Parent'),
				m('CancelRefresh'),
				m('Refresh'),
			]),
		[RANGES]: promotedExcelHostType('Ranges', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', RANGE),
				p('Parent'),
				p('_Default', RANGE),
			]),
		[MODELCONNECTION]: promotedExcelHostType('ModelConnection', [
				p('Application', APPLICATION),
				p('CalculatedMembers', CALCULATEDMEMBERS),
				p('Parent'),
			]),
		[WORKSHEETDATACONNECTION]: promotedExcelHostType('WorksheetDataConnection', [
				p('Application', APPLICATION),
				p('Parent'),
			]),
		[TEXTCONNECTION]: promotedExcelHostType('TextConnection', [
				p('Application', APPLICATION),
				p('Parent'),
			]),
		[DATAFEEDCONNECTION]: promotedExcelHostType('DataFeedConnection', [
				p('Application', APPLICATION),
				p('Parent'),
				m('CancelRefresh'),
				m('Refresh'),
			]),
		[MODELTABLES]: promotedExcelHostType('ModelTables', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', MODELTABLE),
				m('Item', MODELTABLE),
			]),
		[MODELTABLE]: promotedExcelHostType('ModelTable', [
				p('Application', APPLICATION),
				p('ModelTableColumns', MODELTABLECOLUMNS),
				p('Parent'),
				p('SourceWorkbookConnection', WORKBOOKCONNECTION),
				m('Refresh'),
			]),
		[CALCULATEDMEMBER]: promotedExcelHostType('CalculatedMember', [
				p('Application', APPLICATION),
				p('Parent'),
				m('Delete'),
			]),
		[CALCULATEDMEMBERS]: promotedExcelHostType('CalculatedMembers', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', CALCULATEDMEMBER),
				m('Item', CALCULATEDMEMBER),
			]),
		[MODELTABLECOLUMN]: promotedExcelHostType('ModelTableColumn', [
				p('Application', APPLICATION),
				p('Parent'),
			]),
		[MODELTABLECOLUMNS]: promotedExcelHostType('ModelTableColumns', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', MODELTABLECOLUMN),
				m('Item', MODELTABLECOLUMN),
			]),
		[SLICERCACHES]: promotedExcelHostType('SlicerCaches', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent', WORKBOOK),
				p('_Default', SLICERCACHE),
				m('Add', SLICERCACHE),
				m('Add2', SLICERCACHE),
				m('Item', SLICERCACHE),
			]),
		[SLICERCACHE]: promotedExcelHostType('SlicerCache', [
				p('Application', APPLICATION),
				p('Index'),
				p('ListObject', LISTOBJECT),
				p('Parent', WORKBOOK),
				p('PivotTables', SLICERPIVOTTABLES),
				p('SlicerCacheLevels', SLICERCACHELEVELS),
				p('SlicerItems', SLICERITEMS),
				p('Slicers', SLICERS),
				p('TimelineState', TIMELINESTATE),
				p('VisibleSlicerItems', SLICERITEMS),
				p('WorkbookConnection', WORKBOOKCONNECTION),
				m('ClearManualFilter'),
				m('Delete'),
			]),
		[SLICERS]: promotedExcelHostType('Slicers', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', SLICER),
				m('Add', SLICER),
				m('Item', SLICER),
			]),
		[SLICER]: promotedExcelHostType('Slicer', [
				p('Application', APPLICATION),
				p('ActiveItem', SLICERITEM),
				p('Caption'),
				p('Parent'),
				p('Shape', SHAPE),
				p('SlicerCache', SLICERCACHE),
				p('SlicerCacheLevel', SLICERCACHELEVEL),
				p('TimelineViewState', TIMELINEVIEWSTATE),
				m('Delete'),
			]),
		[SLICERITEMS]: promotedExcelHostType('SlicerItems', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', SLICERITEM),
				m('Item', SLICERITEM),
			]),
		[SLICERITEM]: promotedExcelHostType('SlicerItem', [
				p('Application', APPLICATION),
				p('Caption'),
				p('Name'),
				p('Parent'),
				p('Selected'),
				p('SourceName'),
			]),
		[SLICERCACHELEVELS]: promotedExcelHostType('SlicerCacheLevels', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', SLICERCACHELEVEL),
				m('Item', SLICERCACHELEVEL),
			]),
		[SLICERCACHELEVEL]: promotedExcelHostType('SlicerCacheLevel', [
				p('Application', APPLICATION),
				p('Count'),
				p('Name'),
				p('Ordinal'),
				p('Parent'),
				p('SlicerItems', SLICERITEMS),
			]),
		[SLICERPIVOTTABLES]: promotedExcelHostType('SlicerPivotTables', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', PIVOTTABLE),
				m('Item', PIVOTTABLE),
			]),
		[TIMELINESTATE]: promotedExcelHostType('TimelineState', [
				p('Application', APPLICATION),
				p('EndDate'),
				p('Parent'),
				p('StartDate'),
				m('SetFilterDateRange'),
				m('SetFilterDateRange2'),
			]),
		[TIMELINEVIEWSTATE]: promotedExcelHostType('TimelineViewState', [
				p('Application', APPLICATION),
				p('Level'),
				p('Parent'),
				p('ShowHeader'),
				p('ShowHorizontalScrollbar'),
				p('ShowSelectionLabel'),
				p('ShowTimeLevel'),
				p('Visible'),
			]),
		[CHART]: promotedExcelHostType('Chart', [
				p('Application', APPLICATION),
				p('ChartArea', CHARTAREA),
				p('ChartTitle', CHARTTITLE),
				p('ChartType'),
				p('DataTable', DATATABLE),
				p('Floor', FLOOR),
				p('HasLegend'),
				p('HasTitle'),
				p('Legend', LEGEND),
				p('Name'),
				p('PageSetup', PAGESETUP),
				p('Parent'),
				p('PlotArea', PLOTAREA),
				p('Shapes', SHAPES),
				p('Type'),
				p('Visible'),
				p('Walls', WALLS),
				m('Activate'),
				m('ApplyDataLabels'),
				m('Axes', AXES),
				m('ChartGroups', CHARTGROUPS),
				m('ChartObjects', CHARTOBJECTS),
				m('ChartWizard'),
				m('Copy'),
				m('Delete'),
				m('Export'),
				m('FullSeriesCollection', FULLSERIESCOLLECTION),
				m('Location'),
				m('Paste'),
				m('Refresh'),
				m('Select'),
				m('SeriesCollection', SERIESCOLLECTION),
				m('SetSourceData'),
			]),
		[CHARTS]: promotedExcelHostType('Charts', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', CHART),
				p('Parent'),
				p('Visible'),
				m('Add2', CHART),
				m('Copy'),
				m('Delete'),
				m('Select'),
			]),
		[CHARTOBJECT]: promotedExcelHostType('ChartObject', [
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
			]),
		[CHARTOBJECTS]: promotedExcelHostType('ChartObjects', [
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
			]),
		[SERIESCOLLECTION]: promotedExcelHostType('SeriesCollection', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', SERIES),
				m('Item', SERIES),
				m('NewSeries', SERIES),
				m('_Default', SERIES),
			]),
		[FULLSERIESCOLLECTION]: promotedExcelHostType('FullSeriesCollection', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', SERIES),
				m('_Default', SERIES),
			]),
		[SERIES]: promotedExcelHostType('Series', [
				p('Application', APPLICATION),
				p('ErrorBars', ERRORBARS),
				p('Format', CHARTFORMAT),
				p('LeaderLines', LEADERLINES),
				p('Parent'),
				m('DataLabels', DATALABELS),
				m('Points', POINTS),
				m('Trendlines', TRENDLINES),
			]),
		[AXES]: promotedExcelHostType('Axes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', AXIS),
				m('_Default', AXIS),
			]),
		[AXIS]: promotedExcelHostType('Axis', [
				p('Application', APPLICATION),
				p('AxisTitle', AXISTITLE),
				p('Border', BORDER),
				p('DisplayUnitLabel', DISPLAYUNITLABEL),
				p('Format', CHARTFORMAT),
				p('MajorGridlines', GRIDLINES),
				p('MinorGridlines', GRIDLINES),
				p('Parent'),
				p('TickLabels', TICKLABELS),
			]),
		[AXISTITLE]: promotedExcelHostType('AxisTitle', [
				p('Application', APPLICATION),
				p('Characters'),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[CHARTFORMAT]: promotedExcelHostType('ChartFormat', [
				p('Application', APPLICATION),
				p('Fill', FILLFORMAT),
				p('Line', LINEFORMAT),
				p('Parent'),
			]),
		[TICKLABELS]: promotedExcelHostType('TickLabels', [
				p('Application', APPLICATION),
				p('Font', FONT),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[GRIDLINES]: promotedExcelHostType('Gridlines', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[DISPLAYUNITLABEL]: promotedExcelHostType('DisplayUnitLabel', [
				p('Application', APPLICATION),
				p('Characters'),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[LEADERLINES]: promotedExcelHostType('LeaderLines', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[ERRORBARS]: promotedExcelHostType('ErrorBars', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[POINTS]: promotedExcelHostType('Points', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Item', POINT),
				m('_Default', POINT),
			]),
		[POINT]: promotedExcelHostType('Point', [
				p('Application', APPLICATION),
				p('DataLabel', DATALABEL),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[TRENDLINES]: promotedExcelHostType('Trendlines', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				m('Add', TRENDLINE),
				m('Item', TRENDLINE),
				m('_Default', TRENDLINE),
			]),
		[TRENDLINE]: promotedExcelHostType('Trendline', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('DataLabel', DATALABEL),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[DATALABELS]: promotedExcelHostType('DataLabels', [
				p('Application', APPLICATION),
				p('Count'),
				p('Format', CHARTFORMAT),
				p('Parent'),
				m('Item', DATALABEL),
				m('_Default', DATALABEL),
			]),
		[DATALABEL]: promotedExcelHostType('DataLabel', [
				p('Application', APPLICATION),
				p('Characters'),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[CHARTTITLE]: promotedExcelHostType('ChartTitle', [
				p('Application', APPLICATION),
				p('Characters'),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[CHARTAREA]: promotedExcelHostType('ChartArea', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Font', FONT),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[PLOTAREA]: promotedExcelHostType('PlotArea', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[LEGEND]: promotedExcelHostType('Legend', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Font', FONT),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
				m('LegendEntries', LEGENDENTRIES),
			]),
		[LEGENDENTRIES]: promotedExcelHostType('LegendEntries', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', LEGENDENTRY),
				m('Item', LEGENDENTRY),
			]),
		[LEGENDENTRY]: promotedExcelHostType('LegendEntry', [
				p('Application', APPLICATION),
				p('Font', FONT),
				p('Format', CHARTFORMAT),
				p('LegendKey', LEGENDKEY),
				p('Parent'),
			]),
		[LEGENDKEY]: promotedExcelHostType('LegendKey', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[CHARTGROUPS]: promotedExcelHostType('ChartGroups', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', CHARTGROUP),
				m('Item', CHARTGROUP),
			]),
		[CHARTGROUP]: promotedExcelHostType('ChartGroup', [
				p('Application', APPLICATION),
				p('DownBars', DOWNBARS),
				p('DropLines', DROPLINES),
				p('HiLoLines', HILOLINES),
				p('Parent'),
				p('RadarAxisLabels', TICKLABELS),
				p('SeriesLines', SERIESLINES),
				p('UpBars', UPBARS),
				m('FullCategoryCollection'),
				m('SeriesCollection', SERIESCOLLECTION),
			]),
		[DATATABLE]: promotedExcelHostType('DataTable', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Font', FONT),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[WALLS]: promotedExcelHostType('Walls', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[FLOOR]: promotedExcelHostType('Floor', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[SERIESLINES]: promotedExcelHostType('SeriesLines', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[DOWNBARS]: promotedExcelHostType('DownBars', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[DROPLINES]: promotedExcelHostType('DropLines', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[HILOLINES]: promotedExcelHostType('HiLoLines', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Parent'),
			]),
		[UPBARS]: promotedExcelHostType('UpBars', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Format', CHARTFORMAT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[SHAPE]: promotedExcelHostType('Shape', [
				p('AlternativeText'),
				p('Application', APPLICATION),
				p('AutoShapeType'),
				p('BottomRightCell', RANGE),
				p('Callout', CALLOUTFORMAT),
				p('Chart', CHART),
				p('ConnectorFormat', CONNECTORFORMAT),
				p('Fill', FILLFORMAT),
				p('GroupItems', GROUPSHAPES),
				p('Height'),
				p('Left'),
				p('Line', LINEFORMAT),
				p('LockAspectRatio'),
				p('Name'),
				p('Nodes', SHAPENODES),
				p('Parent'),
				p('ParentGroup', SHAPE),
				p('Placement'),
				p('PictureFormat', PICTUREFORMAT),
				p('Rotation'),
				p('Shadow', SHADOWFORMAT),
				p('TextFrame', TEXTFRAME),
				p('TextFrame2', TEXTFRAME2),
				p('ThreeD', THREEDFORMAT),
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
				m('Duplicate', SHAPERANGE),
				m('Flip'),
				m('IncrementLeft'),
				m('IncrementRotation'),
				m('IncrementTop'),
				m('ScaleHeight'),
				m('ScaleWidth'),
				m('Select'),
				m('ZOrder'),
			]),
		[SHAPES]: promotedExcelHostType('Shapes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('Range', SHAPERANGE),
				m('AddChart2', SHAPE),
				m('AddConnector', SHAPE),
				m('AddLabel', SHAPE),
				m('AddLine', SHAPE),
				m('AddPicture', SHAPE),
				m('AddShape', SHAPE),
				m('AddTextbox', SHAPE),
				m('Item', SHAPE),
				m('Range', SHAPERANGE),
				m('SelectAll'),
			]),
		[SHAPERANGE]: promotedExcelHostType('ShapeRange', [
				p('Application', APPLICATION),
				p('Chart', CHART),
				p('Count'),
				p('Fill', FILLFORMAT),
				p('GroupItems', GROUPSHAPES),
				p('Item', SHAPE),
				p('Line', LINEFORMAT),
				p('Nodes', SHAPENODES),
				p('Parent'),
				p('ParentGroup', SHAPE),
				p('PictureFormat', PICTUREFORMAT),
				p('Shadow', SHADOWFORMAT),
				p('TextFrame', TEXTFRAME),
				p('TextFrame2', TEXTFRAME2),
				p('ThreeD', THREEDFORMAT),
				m('Duplicate', SHAPERANGE),
				m('Group', SHAPE),
				m('Item', SHAPE),
				m('Regroup', SHAPE),
				m('Ungroup', SHAPERANGE),
			]),
		[GROUPSHAPES]: promotedExcelHostType('GroupShapes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('Range', SHAPERANGE),
				m('Item', SHAPE),
				m('_Default', SHAPE),
			]),
		[FILLFORMAT]: promotedExcelHostType('FillFormat', [
				p('Application', APPLICATION),
				p('BackColor', COLORFORMAT),
				p('ForeColor', COLORFORMAT),
				p('GradientColorType'),
				p('Parent'),
				p('Pattern'),
				p('PresetGradientType'),
				p('PresetTexture'),
				p('Transparency'),
				p('Type'),
				p('Visible'),
			]),
		[LINEFORMAT]: promotedExcelHostType('LineFormat', [
				p('Application', APPLICATION),
				p('BackColor', COLORFORMAT),
				p('BeginArrowheadStyle'),
				p('EndArrowheadStyle'),
				p('ForeColor', COLORFORMAT),
				p('Parent'),
				p('Pattern'),
				p('Transparency'),
				p('Visible'),
				p('Weight'),
			]),
		[TEXTFRAME]: promotedExcelHostType('TextFrame', [
				p('Application', APPLICATION),
				p('Characters'),
				p('HasText'),
				p('HorizontalAlignment'),
				p('MarginBottom'),
				p('MarginLeft'),
				p('MarginRight'),
				p('MarginTop'),
				p('Orientation'),
				p('Parent'),
				p('VerticalAlignment'),
			]),
		[TEXTFRAME2]: promotedExcelHostType('TextFrame2', [
				p('Application', APPLICATION),
				p('Column'),
				p('HasText'),
				p('MarginBottom'),
				p('MarginLeft'),
				p('MarginRight'),
				p('MarginTop'),
				p('Orientation'),
				p('Parent'),
				p('TextRange'),
				p('ThreeD', THREEDFORMAT),
				p('VerticalAnchor'),
			]),
		[PICTUREFORMAT]: promotedExcelHostType('PictureFormat', [
				p('Application', APPLICATION),
				p('Brightness'),
				p('ColorType'),
				p('Contrast'),
				p('CropBottom'),
				p('CropLeft'),
				p('CropRight'),
				p('CropTop'),
				p('Parent'),
				p('TransparencyColor'),
				p('TransparentBackground'),
			]),
		[SHADOWFORMAT]: promotedExcelHostType('ShadowFormat', [
				p('Application', APPLICATION),
				p('ForeColor', COLORFORMAT),
				p('Obscured'),
				p('OffsetX'),
				p('OffsetY'),
				p('Parent'),
				p('Transparency'),
				p('Type'),
				p('Visible'),
			]),
		[THREEDFORMAT]: promotedExcelHostType('ThreeDFormat', [
				p('Application', APPLICATION),
				p('BevelBottomDepth'),
				p('BevelBottomType'),
				p('BevelTopDepth'),
				p('BevelTopType'),
				p('ContourColor', COLORFORMAT),
				p('ContourWidth'),
				p('Depth'),
				p('ExtrusionColor', COLORFORMAT),
				p('Parent'),
				p('PresetExtrusionDirection'),
				p('PresetLightingDirection'),
				p('PresetLightingSoftness'),
				p('PresetMaterial'),
				p('PresetThreeDFormat'),
				p('RotationX'),
				p('RotationY'),
				p('RotationZ'),
				p('Visible'),
			]),
		[CONNECTORFORMAT]: promotedExcelHostType('ConnectorFormat', [
				p('Application', APPLICATION),
				p('BeginConnected'),
				p('BeginConnectedShape', SHAPE),
				p('BeginConnectionSite'),
				p('EndConnected'),
				p('EndConnectedShape', SHAPE),
				p('EndConnectionSite'),
				p('Parent'),
				p('Type'),
				m('BeginConnect'),
				m('BeginDisconnect'),
				m('EndConnect'),
				m('EndDisconnect'),
			]),
		[CALLOUTFORMAT]: promotedExcelHostType('CalloutFormat', [
				p('Accent'),
				p('Angle'),
				p('Application', APPLICATION),
				p('AutoAttach'),
				p('AutoLength'),
				p('Border'),
				p('Drop'),
				p('Gap'),
				p('Length'),
				p('Parent'),
				p('Type'),
			]),
		[SHAPENODES]: promotedExcelHostType('ShapeNodes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', SHAPENODE),
				m('Delete'),
				m('Insert'),
				m('Item', SHAPENODE),
				m('SetEditingType'),
				m('SetPosition'),
				m('SetSegmentType'),
			]),
		[SHAPENODE]: promotedExcelHostType('ShapeNode', [
				p('Application', APPLICATION),
				p('EditingType'),
				p('Parent'),
				p('Points'),
				p('SegmentType'),
			]),
		[COLORFORMAT]: promotedExcelHostType('ColorFormat', [
				p('Application', APPLICATION),
				p('Brightness'),
				p('Creator'),
				p('ObjectThemeColor'),
				p('Parent'),
				p('RGB'),
				p('SchemeColor'),
				p('TintAndShade'),
				p('Type'),
			]),
		[SORT]: promotedExcelHostType('Sort', [
				p('Application', APPLICATION),
				p('Parent'),
				p('Rng', RANGE),
				p('SortFields', SORTFIELDS),
				m('Apply'),
				m('SetRange'),
			]),
		[SORTFIELDS]: promotedExcelHostType('SortFields', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', SORTFIELD),
				p('Parent'),
				p('_Default', SORTFIELD),
				m('Add', SORTFIELD),
				m('Add2', SORTFIELD),
				m('Clear'),
			]),
		[SORTFIELD]: promotedExcelHostType('SortField', [
				p('Application', APPLICATION),
				p('Key', RANGE),
				p('Parent'),
				m('Delete'),
			]),
		[AUTOFILTER]: promotedExcelHostType('AutoFilter', [
				p('Application', APPLICATION),
				p('Filters', FILTERS),
				p('Parent'),
				p('Range', RANGE),
				p('Sort', SORT),
				m('ApplyFilter'),
				m('ShowAllData'),
			]),
		[FILTERS]: promotedExcelHostType('Filters', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', FILTER),
				p('Parent'),
				p('_Default', FILTER),
			]),
		[FILTER]: promotedExcelHostType('Filter', [
				p('Application', APPLICATION),
				p('Parent'),
			]),
		[ICON]: promotedExcelHostType('Icon', [
				p('Application', APPLICATION),
				p('Parent', ICONSET),
			]),
		[ICONSET]: promotedExcelHostType('IconSet', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', ICON),
				p('Parent'),
				p('_Default', ICON),
			]),
		[ICONSETS]: promotedExcelHostType('IconSets', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', ICONSET),
				p('Parent'),
				p('_Default', ICONSET),
			]),
		[DRAWINGOBJECTS]: promotedExcelHostType('DrawingObjects', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', DRAWING),
				m('Group', GROUPOBJECT),
				m('Item', DRAWING),
			]),
		[DRAWINGS]: promotedExcelHostType('Drawings', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', DRAWING),
				m('Add', DRAWING),
				m('Group', GROUPOBJECT),
				m('Item', DRAWING),
			]),
		[DRAWING]: promotedExcelHostType('Drawing', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
			]),
		[PICTURES]: promotedExcelHostType('Pictures', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', PICTURE),
				m('Add', PICTURE),
				m('Insert', PICTURE),
				m('Item', PICTURE),
			]),
		[PICTURE]: promotedExcelHostType('Picture', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
			]),
		[LINES]: promotedExcelHostType('Lines', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', LINE),
				m('Add', LINE),
				m('Group', GROUPOBJECT),
				m('Item', LINE),
			]),
		[LINE]: promotedExcelHostType('Line', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
			]),
		[RECTANGLES]: promotedExcelHostType('Rectangles', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', RECTANGLE),
				m('Add', RECTANGLE),
				m('Group', GROUPOBJECT),
				m('Item', RECTANGLE),
			]),
		[RECTANGLE]: promotedExcelHostType('Rectangle', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
			]),
		[OVALS]: promotedExcelHostType('Ovals', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', OVAL),
				m('Add', OVAL),
				m('Group', GROUPOBJECT),
				m('Item', OVAL),
			]),
		[OVAL]: promotedExcelHostType('Oval', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
			]),
		[ARCS]: promotedExcelHostType('Arcs', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('_Default', ARC),
				m('Add', ARC),
				m('Group', GROUPOBJECT),
				m('Item', ARC),
			]),
		[ARC]: promotedExcelHostType('Arc', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
				m('Delete'),
				m('Duplicate'),
				m('Select'),
			]),
		[OLEOBJECT]: promotedExcelHostType('OLEObject', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Interior', INTERIOR),
				p('Object'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[OLEOBJECTS]: promotedExcelHostType('OLEObjects', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', OLEOBJECT),
				m('Item', OLEOBJECT),
			]),
		[BUTTON]: promotedExcelHostType('Button', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Font', FONT),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[BUTTONS]: promotedExcelHostType('Buttons', [
				p('Application', APPLICATION),
				p('Count'),
				p('Font', FONT),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', BUTTON),
				m('Group', GROUPOBJECT),
				m('Item', BUTTON),
			]),
		[CHECKBOX]: promotedExcelHostType('CheckBox', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[CHECKBOXES]: promotedExcelHostType('CheckBoxes', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', CHECKBOX),
				m('Group', GROUPOBJECT),
				m('Item', CHECKBOX),
			]),
		[DROPDOWN]: promotedExcelHostType('DropDown', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[DROPDOWNS]: promotedExcelHostType('DropDowns', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', DROPDOWN),
				m('Group', GROUPOBJECT),
				m('Item', DROPDOWN),
			]),
		[OPTIONBUTTON]: promotedExcelHostType('OptionButton', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('GroupBox', GROUPBOX),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[OPTIONBUTTONS]: promotedExcelHostType('OptionButtons', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('GroupBox', GROUPBOX),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', OPTIONBUTTON),
				m('Group', GROUPOBJECT),
				m('Item', OPTIONBUTTON),
			]),
		[GROUPOBJECT]: promotedExcelHostType('GroupObject', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[GROUPOBJECTS]: promotedExcelHostType('GroupObjects', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Group', GROUPOBJECT),
				m('Item', GROUPOBJECT),
			]),
		[GROUPBOX]: promotedExcelHostType('GroupBox', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[GROUPBOXES]: promotedExcelHostType('GroupBoxes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', GROUPBOX),
				m('Group', GROUPOBJECT),
				m('Item', GROUPBOX),
			]),
		[LABEL]: promotedExcelHostType('Label', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[LABELS]: promotedExcelHostType('Labels', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', LABEL),
				m('Group', GROUPOBJECT),
				m('Item', LABEL),
			]),
		[LISTBOX]: promotedExcelHostType('ListBox', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[LISTBOXES]: promotedExcelHostType('ListBoxes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', LISTBOX),
				m('Group', GROUPOBJECT),
				m('Item', LISTBOX),
			]),
		[SCROLLBAR]: promotedExcelHostType('ScrollBar', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[SCROLLBARS]: promotedExcelHostType('ScrollBars', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', SCROLLBAR),
				m('Group', GROUPOBJECT),
				m('Item', SCROLLBAR),
			]),
		[SPINNER]: promotedExcelHostType('Spinner', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[SPINNERS]: promotedExcelHostType('Spinners', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', SPINNER),
				m('Group', GROUPOBJECT),
				m('Item', SPINNER),
			]),
		[EDITBOX]: promotedExcelHostType('EditBox', [
				p('Application', APPLICATION),
				p('BottomRightCell', RANGE),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[EDITBOXES]: promotedExcelHostType('EditBoxes', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', EDITBOX),
				m('Group', GROUPOBJECT),
				m('Item', EDITBOX),
			]),
		[TEXTBOX]: promotedExcelHostType('TextBox', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('BottomRightCell', RANGE),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				p('TopLeftCell', RANGE),
			]),
		[TEXTBOXES]: promotedExcelHostType('TextBoxes', [
				p('Application', APPLICATION),
				p('Border', BORDER),
				p('Count'),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				p('ShapeRange', SHAPERANGE),
				m('Add', TEXTBOX),
				m('Group', GROUPOBJECT),
				m('Item', TEXTBOX),
			]),
		[FONT]: promotedExcelHostType('Font', [
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
			]),
		[INTERIOR]: promotedExcelHostType('Interior', [
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
			]),
		[BORDER]: promotedExcelHostType('Border', [
				p('Application', APPLICATION),
				p('Color'),
				p('ColorIndex'),
				p('LineStyle'),
				p('Parent'),
				p('ThemeColor'),
				p('TintAndShade'),
				p('Weight'),
			]),
		[BORDERS]: promotedExcelHostType('Borders', [
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
			]),
		[AREAS]: promotedExcelHostType('Areas', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', RANGE),
				p('Parent'),
			]),
		[HYPERLINK]: promotedExcelHostType('Hyperlink', [
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
			]),
		[HYPERLINKS]: promotedExcelHostType('Hyperlinks', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', HYPERLINK),
				p('Parent'),
				m('Add', HYPERLINK),
				m('Delete'),
			]),
		[FORMATCONDITION]: promotedExcelHostType('FormatCondition', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('Borders', BORDERS),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				m('Delete'),
				m('Modify'),
			]),
		[FORMATCONDITIONS]: promotedExcelHostType('FormatConditions', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', FORMATCONDITION),
				p('Parent'),
				m('Add', FORMATCONDITION),
				m('AddAboveAverage', ABOVEAVERAGE),
				m('AddColorScale', COLORSCALE),
				m('AddDatabar', DATABAR),
				m('AddIconSetCondition', ICONSETCONDITION),
				m('AddTop10', TOP10),
				m('AddUniqueValues', UNIQUEVALUES),
				m('Delete'),
			]),
		[DATABAR]: promotedExcelHostType('Databar', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('BarBorder', DATABARBORDER),
				p('BarColor', FORMATCOLOR),
				p('MaxPoint', CONDITIONVALUE),
				p('MinPoint', CONDITIONVALUE),
				p('NegativeBarFormat', NEGATIVEBARFORMAT),
				p('Parent'),
				m('Delete'),
				m('ModifyAppliesToRange'),
			]),
		[COLORSCALE]: promotedExcelHostType('ColorScale', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('ColorScaleCriteria', COLORSCALECRITERIA),
				p('Parent'),
				m('Delete'),
				m('ModifyAppliesToRange'),
			]),
		[COLORSCALECRITERIA]: promotedExcelHostType('ColorScaleCriteria', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', COLORSCALECRITERION),
				m('Item', COLORSCALECRITERION),
			]),
		[COLORSCALECRITERION]: promotedExcelHostType('ColorScaleCriterion', [
				p('Application', APPLICATION),
				p('FormatColor', FORMATCOLOR),
				p('Parent'),
				p('Type'),
				p('Value'),
			]),
		[ICONSETCONDITION]: promotedExcelHostType('IconSetCondition', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('IconCriteria', ICONCRITERIA),
				p('IconSet', ICONSET),
				p('Parent'),
				m('Delete'),
				m('ModifyAppliesToRange'),
			]),
		[ICONCRITERIA]: promotedExcelHostType('IconCriteria', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', ICONCRITERION),
				m('Item', ICONCRITERION),
			]),
		[ICONCRITERION]: promotedExcelHostType('IconCriterion', [
				p('Application', APPLICATION),
				p('Icon', ICON),
				p('Operator'),
				p('Parent'),
				p('Type'),
				p('Value'),
			]),
		[ABOVEAVERAGE]: promotedExcelHostType('AboveAverage', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('Borders', BORDERS),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				m('Delete'),
				m('ModifyAppliesToRange'),
			]),
		[TOP10]: promotedExcelHostType('Top10', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('Borders', BORDERS),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				m('Delete'),
				m('ModifyAppliesToRange'),
			]),
		[UNIQUEVALUES]: promotedExcelHostType('UniqueValues', [
				p('Application', APPLICATION),
				p('AppliesTo', RANGE),
				p('Borders', BORDERS),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
				m('Delete'),
				m('ModifyAppliesToRange'),
			]),
		[CELLFORMAT]: promotedExcelHostType('CellFormat', [
				p('Application', APPLICATION),
				p('Borders', BORDERS),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[DISPLAYFORMAT]: promotedExcelHostType('DisplayFormat', [
				p('Application', APPLICATION),
				p('Borders', BORDERS),
				p('Font', FONT),
				p('Interior', INTERIOR),
				p('Parent'),
			]),
		[FORMATCOLOR]: promotedExcelHostType('FormatColor', [
				p('Application', APPLICATION),
				p('Color'),
				p('Parent'),
				p('ThemeColor'),
				p('TintAndShade'),
			]),
		[CONDITIONVALUE]: promotedExcelHostType('ConditionValue', [
				p('Application', APPLICATION),
				p('Parent'),
				p('Type'),
				p('Value'),
			]),
		[DATABARBORDER]: promotedExcelHostType('DataBarBorder', [
				p('Application', APPLICATION),
				p('Color', FORMATCOLOR),
				p('Parent'),
				p('Type'),
			]),
		[NEGATIVEBARFORMAT]: promotedExcelHostType('NegativeBarFormat', [
				p('Application', APPLICATION),
				p('BorderColor', FORMATCOLOR),
				p('Color', FORMATCOLOR),
				p('Parent'),
			]),
		[SPARKLINEGROUPS]: promotedExcelHostType('SparklineGroups', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', SPARKLINEGROUP),
				m('Add', SPARKLINEGROUP),
				m('Item', SPARKLINEGROUP),
			]),
		[SPARKLINEGROUP]: promotedExcelHostType('SparklineGroup', [
				p('Application', APPLICATION),
				p('Axes', SPARKAXES),
				p('Count'),
				p('DateRange', RANGE),
				p('Location', RANGE),
				p('Parent'),
				p('Points', SPARKPOINTS),
				p('SeriesColor', SPARKCOLOR),
				p('SourceData'),
				p('_Default', SPARKLINE),
				m('Delete'),
				m('Item', SPARKLINE),
				m('Modify'),
			]),
		[SPARKLINE]: promotedExcelHostType('Sparkline', [
				p('Application', APPLICATION),
				p('Location', RANGE),
				p('Parent'),
				p('SourceData'),
				m('ModifyLocation'),
				m('ModifySourceData'),
			]),
		[SPARKPOINTS]: promotedExcelHostType('SparkPoints', [
				p('Application', APPLICATION),
				p('Firstpoint', SPARKCOLOR),
				p('Highpoint', SPARKCOLOR),
				p('Lastpoint', SPARKCOLOR),
				p('Lowpoint', SPARKCOLOR),
				p('Markers', SPARKCOLOR),
				p('Negative', SPARKCOLOR),
				p('Parent'),
			]),
		[SPARKCOLOR]: promotedExcelHostType('SparkColor', [
				p('Application', APPLICATION),
				p('Color', FORMATCOLOR),
				p('Parent'),
				p('Visible'),
			]),
		[SPARKAXES]: promotedExcelHostType('SparkAxes', [
				p('Application', APPLICATION),
				p('Horizontal', SPARKHORIZONTALAXIS),
				p('Parent'),
				p('Vertical', SPARKVERTICALAXIS),
			]),
		[SPARKHORIZONTALAXIS]: promotedExcelHostType('SparkHorizontalAxis', [
				p('Application', APPLICATION),
				p('Axis', SPARKCOLOR),
				p('Parent'),
				p('Visible'),
			]),
		[SPARKVERTICALAXIS]: promotedExcelHostType('SparkVerticalAxis', [
				p('Application', APPLICATION),
				p('CustomMaxScaleValue'),
				p('CustomMinScaleValue'),
				p('MaxScaleType'),
				p('MinScaleType'),
				p('Parent'),
			]),
		[XMLMAPS]: promotedExcelHostType('XmlMaps', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', XMLMAP),
				m('Add', XMLMAP),
				m('Item', XMLMAP),
			]),
		[XMLMAP]: promotedExcelHostType('XmlMap', [
				p('Application', APPLICATION),
				p('DataBinding', XMLDATABINDING),
				p('Name'),
				p('Parent'),
				p('RootElementName'),
				p('Schemas', XMLSCHEMAS),
				p('WorkbookConnection', WORKBOOKCONNECTION),
				m('Delete'),
				m('Export'),
				m('Import'),
			]),
		[XPATH]: promotedExcelHostType('XPath', [
				p('Application', APPLICATION),
				p('Map', XMLMAP),
				p('Parent'),
				p('Repeating'),
				p('Value'),
				m('Clear'),
				m('SetValue'),
			]),
		[PUBLISHOBJECTS]: promotedExcelHostType('PublishObjects', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', PUBLISHOBJECT),
				m('Add', PUBLISHOBJECT),
				m('Item', PUBLISHOBJECT),
			]),
		[PUBLISHOBJECT]: promotedExcelHostType('PublishObject', [
				p('Application', APPLICATION),
				p('AutoRepublish'),
				p('Creator'),
				p('DivID'),
				p('Filename'),
				p('HtmlType'),
				p('Parent'),
				p('Sheet'),
				p('Source'),
				p('SourceType'),
				p('Title'),
				m('Delete'),
				m('Publish'),
			]),
		[WEBOPTIONS]: promotedExcelHostType('WebOptions', [
				p('AllowPNG'),
				p('AlwaysSaveInDefaultEncoding'),
				p('Application', APPLICATION),
				p('Encoding'),
				p('FolderSuffix'),
				p('LocationOfComponents'),
				p('Parent'),
				p('PixelsPerInch'),
				p('RelyOnCSS'),
				p('RelyOnVML'),
				p('ScreenSize'),
				p('TargetBrowser'),
				p('UseLongFileNames'),
			]),
		[DEFAULTWEBOPTIONS]: promotedExcelHostType('DefaultWebOptions', [
				p('AllowPNG'),
				p('AlwaysSaveInDefaultEncoding'),
				p('Application', APPLICATION),
				p('Encoding'),
				p('FolderSuffix'),
				p('LocationOfComponents'),
				p('Parent'),
				p('PixelsPerInch'),
				p('RelyOnCSS'),
				p('RelyOnVML'),
				p('ScreenSize'),
				p('TargetBrowser'),
				p('UseLongFileNames'),
			]),
		[XMLNAMESPACES]: promotedExcelHostType('XmlNamespaces', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', XMLNAMESPACE),
				m('Item', XMLNAMESPACE),
			]),
		[XMLNAMESPACE]: promotedExcelHostType('XmlNamespace', [
				p('Application', APPLICATION),
				p('Parent'),
				p('Prefix'),
				p('Uri'),
			]),
		[XMLSCHEMAS]: promotedExcelHostType('XmlSchemas', [
				p('Application', APPLICATION),
				p('Count'),
				p('Parent'),
				p('_Default', XMLSCHEMA),
				m('Item', XMLSCHEMA),
			]),
		[XMLSCHEMA]: promotedExcelHostType('XmlSchema', [
				p('Application', APPLICATION),
				p('Namespace'),
				p('Parent'),
				p('XML'),
			]),
		[XMLDATABINDING]: promotedExcelHostType('XmlDataBinding', [
				p('Application', APPLICATION),
				p('Parent'),
				p('SourceUrl'),
				m('ClearSettings'),
				m('LoadSettings'),
				m('Refresh'),
			]),
		[WORKSHEETFUNCTION]: promotedExcelHostType('WorksheetFunction', [
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
			]),
		[STYLE]: promotedExcelHostType('Style', [
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
			]),
		[STYLES]: promotedExcelHostType('Styles', [
				p('Application', APPLICATION),
				p('Count'),
				p('Item', STYLE),
				p('Parent'),
				m('Add', STYLE),
				m('Merge'),
			]),
		[PAGESETUP]: promotedExcelHostType('PageSetup', [
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
			]),
		[VALIDATION]: promotedExcelHostType('Validation', [
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
			]),
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
});
