import { describe, it, expect } from 'vitest';
import {
	ProjectIndex,
	resolveMemberCompletions,
	resolveMemberDefinitionsAt,
	tokenize,
	resolveHostConstant,
	resolveHostGlobal,
	resolveHostAlias,
	resolveHostMemberSignature,
	resolveMemberReturnType,
	getHostMembers,
	getHostType,
	type HostObjectModel,
} from '../src/analyzer';
import { XLIDE_ASSERT_MODULE_SOURCE } from '../src/vbaTestSupportModule';

/** Offset just after the dot in the first occurrence of `marker` in `src`. */
function dotOffset(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function names(src: string, marker: string, ctx = {}): string[] {
	return resolveMemberCompletions(src, dotOffset(src, marker), ctx).map((m) => m.name);
}

describe('host model resolution', () => {
	it('resolves host globals to qualified types', () => {
		expect(resolveHostGlobal('ThisWorkbook')).toBe('Excel.Workbook');
		expect(resolveHostGlobal('thisworkbook')).toBe('Excel.Workbook');
		expect(resolveHostGlobal('Application')).toBe('Excel.Application');
		expect(resolveHostGlobal('ActiveSheet')).toBe('Excel.Worksheet');
		expect(resolveHostGlobal('ActiveWorkbook')).toBe('Excel.Workbook');
		expect(resolveHostGlobal('NotAGlobal')).toBeUndefined();
	});

	it('resolves generated host enum constants case-insensitively', () => {
		expect(resolveHostConstant('xlUp')?.type).toBe('XlDirection');
		expect(resolveHostConstant('XLUP')?.name).toBe('xlUp');
		expect(resolveHostConstant('msoLineDash')?.type).toBe('MsoLineDashStyle');
		expect(resolveHostConstant('MSOLINEDASH')?.name).toBe('msoLineDash');
		expect(resolveHostConstant('msoLineDash')?.value).toBe(4);
		expect(resolveHostConstant('notAConstant')).toBeUndefined();
	});

	it('resolves As-type aliases case-insensitively', () => {
		expect(resolveHostAlias('Worksheet')).toBe('Excel.Worksheet');
		expect(resolveHostAlias('range')).toBe('Excel.Range');
		expect(resolveHostAlias('Excel.Workbook')).toBe('Excel.Workbook');
		expect(resolveHostAlias('MyClass')).toBeUndefined();
	});

	it('resolves chainable member return types', () => {
		expect(resolveMemberReturnType('Excel.Workbook', 'Worksheets')).toBe('Excel.Worksheets');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Range')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.Range', 'Offset')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.Range', 'Parent')).toBe('Excel.Worksheet');
		expect(resolveMemberReturnType('Excel.Workbook', 'Application')).toBe('Excel.Application');
		expect(resolveMemberReturnType('Excel.Application', 'Worksheets')).toBe('Excel.Worksheets');
		expect(resolveMemberReturnType('Excel.Workbooks', 'Item')).toBe('Excel.Workbook');
		expect(resolveMemberReturnType('Excel.Worksheets', 'Item')).toBe('Excel.Worksheet');
	});

	it('resolves chaining into the broadened object model', () => {
		// Range formatting objects.
		expect(resolveMemberReturnType('Excel.Range', 'Font')).toBe('Excel.Font');
		expect(resolveMemberReturnType('Excel.Range', 'Interior')).toBe('Excel.Interior');
		expect(resolveMemberReturnType('Excel.Range', 'Borders')).toBe('Excel.Borders');
		expect(resolveMemberReturnType('Excel.Borders', 'Item')).toBe('Excel.Border');
		expect(resolveMemberReturnType('Excel.Range', 'FormatConditions')).toBe(
			'Excel.FormatConditions',
		);
		expect(resolveMemberReturnType('Excel.FormatConditions', 'Add')).toBe(
			'Excel.FormatCondition',
		);
		expect(resolveMemberReturnType('Excel.FormatCondition', 'Font')).toBe('Excel.Font');
		expect(resolveMemberReturnType('Excel.Range', 'Validation')).toBe('Excel.Validation');
		expect(resolveMemberReturnType('Excel.Range', 'Hyperlinks')).toBe('Excel.Hyperlinks');
		expect(resolveMemberReturnType('Excel.Hyperlinks', 'Item')).toBe('Excel.Hyperlink');
		expect(resolveMemberReturnType('Excel.Range', 'Areas')).toBe('Excel.Areas');
		// Tables / list objects.
		expect(resolveMemberReturnType('Excel.Worksheet', 'ListObjects')).toBe('Excel.ListObjects');
		expect(resolveMemberReturnType('Excel.ListObjects', 'Add')).toBe('Excel.ListObject');
		expect(resolveMemberReturnType('Excel.ListObject', 'Range')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.ListObject', 'ListColumns')).toBe('Excel.ListColumns');
		expect(resolveMemberReturnType('Excel.ListObject', 'ListRows')).toBe('Excel.ListRows');
		expect(resolveMemberReturnType('Excel.ListColumns', 'Item')).toBe('Excel.ListColumn');
		expect(resolveMemberReturnType('Excel.ListRows', 'Item')).toBe('Excel.ListRow');
		// Windows / names / charts.
		expect(resolveMemberReturnType('Excel.Application', 'ActiveWindow')).toBe('Excel.Window');
		expect(resolveMemberReturnType('Excel.Application', 'Windows')).toBe('Excel.Windows');
		expect(resolveMemberReturnType('Excel.Windows', 'Item')).toBe('Excel.Window');
		expect(resolveMemberReturnType('Excel.Workbook', 'Names')).toBe('Excel.Names');
		expect(resolveMemberReturnType('Excel.Names', 'Add')).toBe('Excel.Name');
		expect(resolveMemberReturnType('Excel.Workbook', 'Charts')).toBe('Excel.Charts');
		expect(resolveMemberReturnType('Excel.Charts', 'Item')).toBe('Excel.Chart');
		expect(resolveMemberReturnType('Excel.Worksheet', 'ChartObjects')).toBe(
			'Excel.ChartObjects',
		);
		expect(resolveMemberReturnType('Excel.ChartObjects', 'Item')).toBe('Excel.ChartObject');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Shapes')).toBe('Excel.Shapes');
		expect(resolveMemberReturnType('Excel.Shapes', 'Item')).toBe('Excel.Shape');
		expect(resolveMemberReturnType('Excel.Worksheet', 'PageSetup')).toBe('Excel.PageSetup');
		// WorksheetFunction is reachable from Application.
		expect(resolveMemberReturnType('Excel.Application', 'WorksheetFunction')).toBe(
			'Excel.WorksheetFunction',
		);
		expect(resolveHostMemberSignature('Excel.WorksheetFunction', 'Sum')).toContain(
			'Sum(Arg1 As Variant',
		);
		// Pivot tables and fields.
		expect(resolveMemberReturnType('Excel.Worksheet', 'PivotTables')).toBe('Excel.PivotTables');
		expect(resolveMemberReturnType('Excel.PivotTables', 'Item')).toBe('Excel.PivotTable');
		expect(resolveMemberReturnType('Excel.Workbook', 'PivotCaches')).toBe('Excel.PivotCaches');
		expect(resolveMemberReturnType('Excel.PivotCaches', 'Item')).toBe('Excel.PivotCache');
		expect(resolveMemberReturnType('Excel.PivotTable', 'PivotFields')).toBe('Excel.PivotFields');
		expect(resolveMemberReturnType('Excel.PivotFields', 'Item')).toBe('Excel.PivotField');
		expect(resolveMemberReturnType('Excel.PivotField', 'PivotItems')).toBe('Excel.PivotItems');
		expect(resolveMemberReturnType('Excel.PivotItems', 'Item')).toBe('Excel.PivotItem');
		expect(resolveMemberReturnType('Excel.PivotTable', 'PivotCache')).toBe('Excel.PivotCache');
		expect(resolveMemberReturnType('Excel.PivotTable', 'PivotFilters')).toBe(
			'Excel.PivotFilters',
		);
		expect(resolveMemberReturnType('Excel.PivotFilters', 'Item')).toBe('Excel.PivotFilter');
		expect(resolveMemberReturnType('Excel.PivotTable', 'CalculatedFields')).toBe(
			'Excel.CalculatedFields',
		);
		expect(resolveMemberReturnType('Excel.CalculatedFields', 'Add')).toBe('Excel.PivotField');
		expect(resolveMemberReturnType('Excel.PivotField', 'CalculatedItems')).toBe(
			'Excel.CalculatedItems',
		);
		expect(resolveMemberReturnType('Excel.CalculatedItems', 'Add')).toBe('Excel.PivotItem');
		expect(resolveMemberReturnType('Excel.PivotTable', 'CubeFields')).toBe('Excel.CubeFields');
		expect(resolveMemberReturnType('Excel.CubeFields', 'Item')).toBe('Excel.CubeField');
		// Query tables / workbook connections.
		expect(resolveMemberReturnType('Excel.Worksheet', 'QueryTables')).toBe('Excel.QueryTables');
		expect(resolveMemberReturnType('Excel.QueryTables', 'Item')).toBe('Excel.QueryTable');
		expect(resolveMemberReturnType('Excel.QueryTable', 'ResultRange')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.QueryTable', 'Parameters')).toBe('Excel.Parameters');
		expect(resolveMemberReturnType('Excel.Parameters', 'Item')).toBe('Excel.Parameter');
		expect(resolveMemberReturnType('Excel.QueryTable', 'WorkbookConnection')).toBe(
			'Excel.WorkbookConnection',
		);
		expect(resolveMemberReturnType('Excel.Workbook', 'Connections')).toBe('Excel.Connections');
		expect(resolveMemberReturnType('Excel.Connections', 'Item')).toBe(
			'Excel.WorkbookConnection',
		);
		expect(resolveMemberReturnType('Excel.WorkbookConnection', 'OLEDBConnection')).toBe(
			'Excel.OLEDBConnection',
		);
		expect(resolveMemberReturnType('Excel.WorkbookConnection', 'ModelTables')).toBe(
			'Excel.ModelTables',
		);
		expect(resolveMemberReturnType('Excel.ModelTables', 'Item')).toBe('Excel.ModelTable');
		expect(resolveMemberReturnType('Excel.ModelTable', 'ModelTableColumns')).toBe(
			'Excel.ModelTableColumns',
		);
		// Slicers / timelines.
		expect(resolveMemberReturnType('Excel.Workbook', 'SlicerCaches')).toBe(
			'Excel.SlicerCaches',
		);
		expect(resolveMemberReturnType('Excel.SlicerCaches', 'Item')).toBe('Excel.SlicerCache');
		expect(resolveMemberReturnType('Excel.SlicerCache', 'Slicers')).toBe('Excel.Slicers');
		expect(resolveMemberReturnType('Excel.Slicers', 'Item')).toBe('Excel.Slicer');
		expect(resolveMemberReturnType('Excel.Slicer', 'Shape')).toBe('Excel.Shape');
		expect(resolveMemberReturnType('Excel.SlicerCache', 'SlicerItems')).toBe(
			'Excel.SlicerItems',
		);
		expect(resolveMemberReturnType('Excel.SlicerItems', 'Item')).toBe('Excel.SlicerItem');
		expect(resolveMemberReturnType('Excel.SlicerCache', 'SlicerCacheLevels')).toBe(
			'Excel.SlicerCacheLevels',
		);
		expect(resolveMemberReturnType('Excel.SlicerCacheLevels', 'Item')).toBe(
			'Excel.SlicerCacheLevel',
		);
		expect(resolveMemberReturnType('Excel.SlicerCacheLevel', 'SlicerItems')).toBe(
			'Excel.SlicerItems',
		);
		expect(resolveMemberReturnType('Excel.SlicerCache', 'TimelineState')).toBe(
			'Excel.TimelineState',
		);
		// Chart internals.
		expect(resolveMemberReturnType('Excel.Chart', 'SeriesCollection')).toBe(
			'Excel.SeriesCollection',
		);
		expect(resolveMemberReturnType('Excel.SeriesCollection', 'Item')).toBe('Excel.Series');
		expect(resolveMemberReturnType('Excel.Chart', 'FullSeriesCollection')).toBe(
			'Excel.FullSeriesCollection',
		);
		expect(resolveMemberReturnType('Excel.FullSeriesCollection', 'Item')).toBe(
			'Excel.Series',
		);
		expect(resolveMemberReturnType('Excel.Series', 'Points')).toBe('Excel.Points');
		expect(resolveMemberReturnType('Excel.Points', 'Item')).toBe('Excel.Point');
		expect(resolveMemberReturnType('Excel.Series', 'Trendlines')).toBe('Excel.Trendlines');
		expect(resolveMemberReturnType('Excel.Trendlines', 'Item')).toBe('Excel.Trendline');
		expect(resolveMemberReturnType('Excel.Series', 'DataLabels')).toBe('Excel.DataLabels');
		expect(resolveMemberReturnType('Excel.DataLabels', 'Item')).toBe('Excel.DataLabel');
		expect(resolveMemberReturnType('Excel.Chart', 'Axes')).toBe('Excel.Axes');
		expect(resolveMemberReturnType('Excel.Axes', 'Item')).toBe('Excel.Axis');
		expect(resolveMemberReturnType('Excel.Axis', 'AxisTitle')).toBe('Excel.AxisTitle');
		expect(resolveMemberReturnType('Excel.Axis', 'TickLabels')).toBe('Excel.TickLabels');
		expect(resolveMemberReturnType('Excel.Chart', 'ChartTitle')).toBe('Excel.ChartTitle');
		expect(resolveMemberReturnType('Excel.Chart', 'ChartArea')).toBe('Excel.ChartArea');
		expect(resolveMemberReturnType('Excel.Chart', 'PlotArea')).toBe('Excel.PlotArea');
		expect(resolveMemberReturnType('Excel.Chart', 'Legend')).toBe('Excel.Legend');
		expect(resolveMemberReturnType('Excel.Legend', 'LegendEntries')).toBe(
			'Excel.LegendEntries',
		);
		expect(resolveMemberReturnType('Excel.LegendEntries', 'Item')).toBe('Excel.LegendEntry');
		expect(resolveMemberReturnType('Excel.LegendEntry', 'LegendKey')).toBe('Excel.LegendKey');
		expect(resolveMemberReturnType('Excel.Chart', 'ChartGroups')).toBe('Excel.ChartGroups');
		expect(resolveMemberReturnType('Excel.ChartGroups', 'Item')).toBe('Excel.ChartGroup');
		expect(resolveMemberReturnType('Excel.ChartGroup', 'SeriesLines')).toBe(
			'Excel.SeriesLines',
		);
		expect(resolveMemberReturnType('Excel.Chart', 'DataTable')).toBe('Excel.DataTable');
		expect(resolveMemberReturnType('Excel.Chart', 'Walls')).toBe('Excel.Walls');
		expect(resolveMemberReturnType('Excel.Chart', 'Floor')).toBe('Excel.Floor');
		// Shape ranges / comments / sort-filter / controls.
		expect(resolveMemberReturnType('Excel.Shapes', 'Range')).toBe('Excel.ShapeRange');
		expect(resolveMemberReturnType('Excel.ShapeRange', 'Item')).toBe('Excel.Shape');
		expect(resolveMemberReturnType('Excel.ShapeRange', 'GroupItems')).toBe('Excel.GroupShapes');
		expect(resolveMemberReturnType('Excel.GroupShapes', 'Item')).toBe('Excel.Shape');
		expect(resolveMemberReturnType('Excel.Shape', 'Fill')).toBe('Excel.FillFormat');
		expect(resolveMemberReturnType('Excel.FillFormat', 'ForeColor')).toBe('Excel.ColorFormat');
		expect(resolveMemberReturnType('Excel.Shape', 'Line')).toBe('Excel.LineFormat');
		expect(resolveMemberReturnType('Excel.Shape', 'TextFrame')).toBe('Excel.TextFrame');
		expect(resolveMemberReturnType('Excel.Shape', 'TextFrame2')).toBe('Excel.TextFrame2');
		expect(resolveMemberReturnType('Excel.Shape', 'PictureFormat')).toBe(
			'Excel.PictureFormat',
		);
		expect(resolveMemberReturnType('Excel.Shape', 'Shadow')).toBe('Excel.ShadowFormat');
		expect(resolveMemberReturnType('Excel.Shape', 'ThreeD')).toBe('Excel.ThreeDFormat');
		expect(resolveMemberReturnType('Excel.Shape', 'ConnectorFormat')).toBe(
			'Excel.ConnectorFormat',
		);
		expect(resolveMemberReturnType('Excel.Shape', 'Callout')).toBe('Excel.CalloutFormat');
		expect(resolveMemberReturnType('Excel.Shape', 'Nodes')).toBe('Excel.ShapeNodes');
		expect(resolveMemberReturnType('Excel.ShapeNodes', 'Item')).toBe('Excel.ShapeNode');
		expect(resolveMemberReturnType('Excel.Range', 'CommentThreaded')).toBe(
			'Excel.CommentThreaded',
		);
		expect(resolveMemberReturnType('Excel.CommentThreaded', 'Replies')).toBe(
			'Excel.CommentsThreaded',
		);
		expect(resolveMemberReturnType('Excel.CommentsThreaded', 'Item')).toBe(
			'Excel.CommentThreaded',
		);
		expect(resolveMemberReturnType('Excel.Worksheet', 'AutoFilter')).toBe('Excel.AutoFilter');
		expect(resolveMemberReturnType('Excel.AutoFilter', 'Filters')).toBe('Excel.Filters');
		expect(resolveMemberReturnType('Excel.Filters', 'Item')).toBe('Excel.Filter');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Sort')).toBe('Excel.Sort');
		expect(resolveMemberReturnType('Excel.Sort', 'SortFields')).toBe('Excel.SortFields');
		expect(resolveMemberReturnType('Excel.SortFields', 'Add')).toBe('Excel.SortField');
		expect(resolveMemberReturnType('Excel.Range', 'DisplayFormat')).toBe('Excel.DisplayFormat');
		expect(resolveMemberReturnType('Excel.FormatConditions', 'AddDatabar')).toBe(
			'Excel.Databar',
		);
		expect(resolveMemberReturnType('Excel.Databar', 'MinPoint')).toBe('Excel.ConditionValue');
		expect(resolveMemberReturnType('Excel.Databar', 'BarBorder')).toBe('Excel.DataBarBorder');
		expect(resolveMemberReturnType('Excel.Databar', 'NegativeBarFormat')).toBe(
			'Excel.NegativeBarFormat',
		);
		expect(resolveMemberReturnType('Excel.FormatConditions', 'AddColorScale')).toBe(
			'Excel.ColorScale',
		);
		expect(resolveMemberReturnType('Excel.ColorScale', 'ColorScaleCriteria')).toBe(
			'Excel.ColorScaleCriteria',
		);
		expect(resolveMemberReturnType('Excel.ColorScaleCriteria', 'Item')).toBe(
			'Excel.ColorScaleCriterion',
		);
		expect(resolveMemberReturnType('Excel.ColorScaleCriterion', 'FormatColor')).toBe(
			'Excel.FormatColor',
		);
		expect(resolveMemberReturnType('Excel.FormatConditions', 'AddIconSetCondition')).toBe(
			'Excel.IconSetCondition',
		);
		expect(resolveMemberReturnType('Excel.IconSetCondition', 'IconCriteria')).toBe(
			'Excel.IconCriteria',
		);
		expect(resolveMemberReturnType('Excel.IconCriteria', 'Item')).toBe('Excel.IconCriterion');
		expect(resolveMemberReturnType('Excel.FormatConditions', 'AddTop10')).toBe('Excel.Top10');
		expect(resolveMemberReturnType('Excel.FormatConditions', 'AddAboveAverage')).toBe(
			'Excel.AboveAverage',
		);
		expect(resolveMemberReturnType('Excel.FormatConditions', 'AddUniqueValues')).toBe(
			'Excel.UniqueValues',
		);
		expect(resolveMemberReturnType('Excel.Worksheet', 'Drawings')).toBe('Excel.Drawings');
		expect(resolveMemberReturnType('Excel.Drawings', 'Add')).toBe('Excel.Drawing');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Pictures')).toBe('Excel.Pictures');
		expect(resolveMemberReturnType('Excel.Pictures', 'Item')).toBe('Excel.Picture');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Lines')).toBe('Excel.Lines');
		expect(resolveMemberReturnType('Excel.Lines', 'Item')).toBe('Excel.Line');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Rectangles')).toBe('Excel.Rectangles');
		expect(resolveMemberReturnType('Excel.Rectangles', 'Item')).toBe('Excel.Rectangle');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Ovals')).toBe('Excel.Ovals');
		expect(resolveMemberReturnType('Excel.Ovals', 'Item')).toBe('Excel.Oval');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Arcs')).toBe('Excel.Arcs');
		expect(resolveMemberReturnType('Excel.Arcs', 'Item')).toBe('Excel.Arc');
		expect(resolveMemberReturnType('Excel.Worksheet', 'OLEObjects')).toBe('Excel.OLEObjects');
		expect(resolveMemberReturnType('Excel.OLEObjects', 'Item')).toBe('Excel.OLEObject');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Buttons')).toBe('Excel.Buttons');
		expect(resolveMemberReturnType('Excel.Buttons', 'Item')).toBe('Excel.Button');
		expect(resolveMemberReturnType('Excel.CheckBoxes', 'Item')).toBe('Excel.CheckBox');
		expect(resolveMemberReturnType('Excel.DropDowns', 'Item')).toBe('Excel.DropDown');
		expect(resolveMemberReturnType('Excel.OptionButtons', 'Item')).toBe(
			'Excel.OptionButton',
		);
		expect(resolveMemberReturnType('Excel.Range', 'SparklineGroups')).toBe(
			'Excel.SparklineGroups',
		);
		expect(resolveMemberReturnType('Excel.SparklineGroups', 'Item')).toBe(
			'Excel.SparklineGroup',
		);
		expect(resolveMemberReturnType('Excel.SparklineGroup', 'Item')).toBe('Excel.Sparkline');
		expect(resolveMemberReturnType('Excel.SparklineGroup', 'Points')).toBe(
			'Excel.SparkPoints',
		);
		expect(resolveMemberReturnType('Excel.SparkPoints', 'Highpoint')).toBe('Excel.SparkColor');
		expect(resolveMemberReturnType('Excel.SparkColor', 'Color')).toBe('Excel.FormatColor');
		expect(resolveMemberReturnType('Excel.SparklineGroup', 'Axes')).toBe('Excel.SparkAxes');
		expect(resolveMemberReturnType('Excel.SparkAxes', 'Horizontal')).toBe(
			'Excel.SparkHorizontalAxis',
		);
		expect(resolveMemberReturnType('Excel.Workbook', 'XmlMaps')).toBe('Excel.XmlMaps');
		expect(resolveMemberReturnType('Excel.XmlMaps', 'Item')).toBe('Excel.XmlMap');
		expect(resolveMemberReturnType('Excel.XmlMap', 'WorkbookConnection')).toBe(
			'Excel.WorkbookConnection',
		);
		expect(resolveMemberReturnType('Excel.XmlMap', 'Schemas')).toBe('Excel.XmlSchemas');
		expect(resolveMemberReturnType('Excel.XmlMap', 'DataBinding')).toBe('Excel.XmlDataBinding');
		expect(resolveMemberReturnType('Excel.ListColumn', 'XPath')).toBe('Excel.XPath');
		expect(resolveMemberReturnType('Excel.XPath', 'Map')).toBe('Excel.XmlMap');
		expect(resolveMemberReturnType('Excel.Workbook', 'PublishObjects')).toBe(
			'Excel.PublishObjects',
		);
		expect(resolveMemberReturnType('Excel.PublishObjects', 'Item')).toBe(
			'Excel.PublishObject',
		);
		expect(resolveMemberReturnType('Excel.Workbook', 'WebOptions')).toBe('Excel.WebOptions');
		expect(resolveMemberReturnType('Excel.Application', 'DefaultWebOptions')).toBe(
			'Excel.DefaultWebOptions',
		);
	});

	it('exposes the broadened types as host aliases', () => {
		expect(resolveHostAlias('Font')).toBe('Excel.Font');
		expect(resolveHostAlias('listobject')).toBe('Excel.ListObject');
		expect(resolveHostAlias('Window')).toBe('Excel.Window');
		expect(resolveHostAlias('FormatCondition')).toBe('Excel.FormatCondition');
		expect(resolveHostAlias('FormatConditions')).toBe('Excel.FormatConditions');
		expect(resolveHostAlias('ChartObject')).toBe('Excel.ChartObject');
		expect(resolveHostAlias('Hyperlink')).toBe('Excel.Hyperlink');
		expect(resolveHostAlias('WorksheetFunction')).toBe('Excel.WorksheetFunction');
		expect(resolveHostAlias('PivotTable')).toBe('Excel.PivotTable');
		expect(resolveHostAlias('PivotField')).toBe('Excel.PivotField');
		expect(resolveHostAlias('PivotCache')).toBe('Excel.PivotCache');
		expect(resolveHostAlias('QueryTable')).toBe('Excel.QueryTable');
		expect(resolveHostAlias('SeriesCollection')).toBe('Excel.SeriesCollection');
		expect(resolveHostAlias('ShapeRange')).toBe('Excel.ShapeRange');
		expect(resolveHostAlias('CommentThreaded')).toBe('Excel.CommentThreaded');
		expect(resolveHostAlias('SortFields')).toBe('Excel.SortFields');
		expect(resolveHostAlias('OLEObject')).toBe('Excel.OLEObject');
		expect(resolveHostAlias('WorkbookConnection')).toBe('Excel.WorkbookConnection');
		expect(resolveHostAlias('SlicerCache')).toBe('Excel.SlicerCache');
		expect(resolveHostAlias('FillFormat')).toBe('Excel.FillFormat');
		expect(resolveHostAlias('ChartTitle')).toBe('Excel.ChartTitle');
		expect(resolveHostAlias('Databar')).toBe('Excel.Databar');
		expect(resolveHostAlias('SparklineGroup')).toBe('Excel.SparklineGroup');
		expect(resolveHostAlias('XmlMap')).toBe('Excel.XmlMap');
	});
});

describe('member completion - collections', () => {
	it('offers Workbooks collection members after the Workbooks global', () => {
		const src = 'Sub Test()\n    Workbooks.\nEnd Sub\n';
		const got = names(src, 'Workbooks.');
		expect(got).toContain('Add');
		expect(got).toContain('Open');
		expect(got).toContain('Item');
		expect(got).toContain('Count');
		expect(got).toContain('Close');
	});

	it('resolves the Workbooks global case-insensitively', () => {
		const src = 'Sub Test()\n    workbooks.\nEnd Sub\n';
		const got = names(src, 'workbooks.');
		expect(got).toContain('Add');
		expect(got).toContain('Count');
	});

	it('offers Worksheets collection members after the Worksheets global', () => {
		const src = 'Sub Test()\n    Worksheets.\nEnd Sub\n';
		const got = names(src, 'Worksheets.');
		expect(got).toContain('Add');
		expect(got).toContain('Item');
		expect(got).toContain('Count');
	});

	it('chains a Workbook collection through to its element type', () => {
		// Workbooks.Item returns a Workbook -> Workbook members follow.
		const src = 'Sub Test()\n    Workbooks.Item(1).\nEnd Sub\n';
		const got = names(src, 'Workbooks.Item(1).');
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
	});

	it('chains ThisWorkbook.Worksheets to the Worksheets collection', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Worksheets.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.Worksheets.');
		expect(got).toContain('Add');
		expect(got).toContain('Count');
	});

	it('keeps Sheets.Item ambiguous for single-return queries', () => {
		expect(resolveMemberReturnType('Excel.Sheets', 'Item')).toBeUndefined();
	});

	it('excludes host events from object member surfaces', () => {
		const workbook = getHostMembers('Excel.Workbook').map((member) => member.name);
		const worksheet = getHostMembers('Excel.Worksheet').map((member) => member.name);
		const application = getHostMembers('Excel.Application').map((member) => member.name);
		expect(workbook).not.toContain('AfterSave');
		expect(workbook).not.toContain('Open');
		expect(worksheet).not.toContain('Change');
		expect(application).not.toContain('SheetCalculate');
		expect(resolveHostMemberSignature('Excel.Workbook', 'AfterSave')).toBeUndefined();
		expect(resolveMemberReturnType('Excel.Workbook', 'AfterSave')).toBeUndefined();

		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					members: [{ name: 'Changed', kind: 'event', signature: 'Changed()' }],
				},
			},
			memberSignatures: { 'Test.Thing': { changed: 'Changed()' } },
		};
		expect(getHostMembers('Test.Thing', model)).toEqual([]);
		expect(resolveHostMemberSignature('Test.Thing', 'Changed', model)).toBeUndefined();
	});
});

describe('member completion - host globals', () => {
	it('offers verified Workbook members after ThisWorkbook.', () => {
		const src = 'Sub Test()\n    ThisWorkbook.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.');
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
		expect(got).toContain('SaveAs');
		expect(got).toContain('Close');
		expect(got).toContain('Name');
		expect(got).toContain('FullName');
		expect(got).toContain('VBProject');
	});

	it('offers Application members after Application.', () => {
		const src = 'Sub Test()\n    Application.\nEnd Sub\n';
		const got = names(src, 'Application.');
		expect(got).toContain('ScreenUpdating');
		expect(got).toContain('Workbooks');
		expect(got).toContain('ThisWorkbook');
		expect(got).toContain('Calculate');
	});

	it('offers Worksheet members after ActiveSheet.', () => {
		const src = 'Sub Test()\n    ActiveSheet.\nEnd Sub\n';
		const got = names(src, 'ActiveSheet.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).toContain('Activate');
	});

	it('filters by the partial member prefix already typed', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Sav\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.Sav');
		expect(got).toContain('Save');
		expect(got).toContain('SaveAs');
		expect(got).toContain('SaveCopyAs');
		expect(got).not.toContain('Close');
	});

	it('does not invent members', () => {
		const src = 'Sub Test()\n    ThisWorkbook.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.');
		expect(got).not.toContain('FooBar');
		expect(got).not.toContain('DoesNotExist');
	});

	it('includes dump-backed Workbook members after ThisWorkbook.', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Accept\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ThisWorkbook.Accept'));
		const accept = got.find((member) => member.name === 'AcceptAllChanges');
		expect(accept?.owner).toBe('Excel.Workbook');
		expect(accept?.surfaceExhaustive).toBe(true);
	});

	it('includes generated members on promoted host surfaces with per-type exhaustiveness', () => {
		const appSrc = 'Sub Test()\n    Application.Centi\nEnd Sub\n';
		const app = resolveMemberCompletions(appSrc, dotOffset(appSrc, 'Application.Centi'));
		const centimetersToPoints = app.find((member) => member.name === 'CentimetersToPoints');
		expect(centimetersToPoints?.owner).toBe('Excel.Application');
		expect(centimetersToPoints?.surfaceExhaustive).toBe(true);

		const rangeSrc = 'Sub Test(rng As Range)\n    rng.Spilling\nEnd Sub\n';
		const range = resolveMemberCompletions(rangeSrc, dotOffset(rangeSrc, 'rng.Spilling'));
		const spillingToRange = range.find((member) => member.name === 'SpillingToRange');
		expect(spillingToRange?.owner).toBe('Excel.Range');
		expect(spillingToRange?.surfaceExhaustive).toBe(true);

		const sheetSrc = 'Sub Test(ws As Worksheet)\n    ws.Named\nEnd Sub\n';
		const sheet = resolveMemberCompletions(sheetSrc, dotOffset(sheetSrc, 'ws.Named'));
		const namedSheetViews = sheet.find((member) => member.name === 'NamedSheetViews');
		expect(namedSheetViews?.owner).toBe('Excel.Worksheet');
		expect(namedSheetViews?.surfaceExhaustive).toBe(true);
	});

	it('does not offer Excel events as object member completions', () => {
		expect(names('Sub Test()\n    ThisWorkbook.After\nEnd Sub\n', 'ThisWorkbook.After')).not.toContain(
			'AfterSave',
		);
		expect(names('Sub Test()\n    Application.SheetCalc\nEnd Sub\n', 'Application.SheetCalc')).not.toContain(
			'SheetCalculate',
		);
		expect(names('Sub Test()\n    ActiveSheet.Change\nEnd Sub\n', 'ActiveSheet.Change')).not.toContain(
			'Change',
		);
	});

	it('uses the dump-backed Workbook surface for ActiveWorkbook and Workbook variables', () => {
		const activeSrc = 'Sub Test()\n    ActiveWorkbook.Accept\nEnd Sub\n';
		const active = resolveMemberCompletions(
			activeSrc,
			dotOffset(activeSrc, 'ActiveWorkbook.Accept'),
		);
		expect(
			active.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive,
		).toBe(true);

		const variableSrc =
			'Sub Test()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.Accept\n' +
			'End Sub\n';
		const variable = resolveMemberCompletions(
			variableSrc,
			dotOffset(variableSrc, 'wb.Accept'),
		);
		expect(
			variable.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive,
		).toBe(true);
	});
});

describe('member completion - runtime objects', () => {
	it('offers verified Debug object members', () => {
		const src = 'Sub Test()\n    Debug.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Debug.'));
		const print = got.find((member) => member.name === 'Print');
		const assert = got.find((member) => member.name === 'Assert');

		expect(print?.owner).toBe('Debug');
		expect(print?.kind).toBe('method');
		expect(print?.signature).toContain('ParamArray OutputList As Variant');
		expect(print?.surfaceExhaustive).toBe(true);
		expect(assert?.signature).toContain('BooleanExpression As Boolean');
	});

	it('offers verified Err object members', () => {
		const src = 'Sub Test()\n    Err.Rai\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Err.Rai'));
		const raise = got.find((member) => member.name === 'Raise');
		expect(raise?.owner).toBe('Err');
		expect(raise?.kind).toBe('method');
		expect(raise?.signature).toContain('Number As Long');
		expect(raise?.surfaceExhaustive).toBe(true);
	});

	it('does not treat scalar Err object properties as chainable objects', () => {
		const src = 'Sub Test()\n    Err.Number.\nEnd Sub\n';
		expect(resolveMemberCompletions(src, dotOffset(src, 'Err.Number.'))).toEqual([]);
	});
});

describe('member completion - standard modules', () => {
	it('offers installed XLIDE assertion helpers after XlideAssert.', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'XlideAssert',
			moduleKind: 'standard',
			source: XLIDE_ASSERT_MODULE_SOURCE,
		});
		index.setModule({ moduleName: 'Tests', moduleKind: 'standard', source: '' });
		const src = 'Sub TestInvoice()\n    XlideAssert.Are\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'XlideAssert.Are'), {
			projectClassMembers: index.projectMemberSurfaces('Tests'),
		});
		const areEqual = got.find((member) => member.name === 'AreEqual');

		expect(got.map((member) => member.name)).toContain('AreNotEqual');
		expect(areEqual?.owner).toBe('XlideAssert');
		expect(areEqual?.kind).toBe('method');
		expect(areEqual?.signature).toContain('AreEqual(expected As Variant, actual As Variant, [message As String = ""])');
		expect(areEqual?.surfaceExhaustive).toBe(true);

		const containsSrc = 'Sub TestInvoice()\n    XlideAssert.Con\nEnd Sub\n';
		const contains = resolveMemberCompletions(containsSrc, dotOffset(containsSrc, 'XlideAssert.Con'), {
			projectClassMembers: index.projectMemberSurfaces('Tests'),
		}).find((member) => member.name === 'Contains');
		expect(contains?.signature).toContain(
			'Contains(actual As Variant, expectedSubstring As Variant, [message As String = ""])',
		);
	});

	it('offers exported procedures, constants, enums, enum members, and declares after ModuleName.', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Finance',
			moduleKind: 'standard',
			source: [
				'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency',
				'End Function',
				'Public Const DefaultTaxRate As Double = 0.08',
				'Public Enum SharedMode',
				'    SharedOnly',
				'End Enum',
				'Private Const HiddenTaxRate As Double = 0.99',
				'Private Enum HiddenMode',
				'    HiddenOnly',
				'End Enum',
				'Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
			].join('\n'),
		});
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });

		const src = 'Sub TestInvoice()\n    Finance.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Finance.'), {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		const names = got.map((member) => member.name);
		const invoiceTotal = got.find((member) => member.name === 'InvoiceTotal');
		const defaultTaxRate = got.find((member) => member.name === 'DefaultTaxRate');
		const sharedOnly = got.find((member) => member.name === 'SharedOnly');
		const getTickCount = got.find((member) => member.name === 'GetTickCount');

		expect(names).toEqual([
			'InvoiceTotal',
			'DefaultTaxRate',
			'SharedMode',
			'SharedOnly',
			'GetTickCount',
		]);
		expect(invoiceTotal?.kind).toBe('method');
		expect(invoiceTotal?.signature).toBe('InvoiceTotal(Subtotal As Currency) As Currency');
		expect(defaultTaxRate).toMatchObject({
			kind: 'property',
			returns: 'Double',
			writable: false,
			owner: 'Finance',
		});
		expect(sharedOnly).toMatchObject({
			kind: 'property',
			returns: 'SharedMode',
			writable: false,
		});
		expect(getTickCount?.kind).toBe('method');
		expect(getTickCount?.signature).toBe('GetTickCount() As Long');
		expect(names).not.toContain('HiddenTaxRate');
		expect(names).not.toContain('HiddenOnly');
	});
});

describe('member completion - code names and Me', () => {
	it('resolves a worksheet code name from project context', () => {
		const src = 'Sub Test()\n    Sheet1.\nEnd Sub\n';
		const ctx = { codeNames: { sheet1: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet1.', ctx);
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
	});

	it('resolves a chart sheet code name from project context', () => {
		const src = 'Sub Test()\n    Chart1.\nEnd Sub\n';
		const ctx = { codeNames: { chart1: 'Excel.Chart' } };
		const got = names(src, 'Chart1.', ctx);
		expect(got).toContain('ChartType');
		expect(got).toContain('SeriesCollection');
		expect(got).not.toContain('Range');
	});

	it('does not merge a dangling dot on a previous line into the chain', () => {
		// `wb.` on its own line must not combine with `Sheet3.` below it.
		const src =
			'Sub Test()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.\n' +
			'\n' +
			'    Sheet3.\n' +
			'End Sub\n';
		const ctx = { codeNames: { sheet3: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet3.', ctx);
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
	});

	it('terminates the receiver chain at a colon separator', () => {
		const src = 'Sub Test()\n    Dim wb As Workbook\n    wb.Save : Sheet3.\nEnd Sub\n';
		const ctx = { codeNames: { sheet3: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet3.', ctx);
		expect(got).toContain('Range');
	});

	it('resolves Me to the module host type (Worksheet)', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = names(src, 'Me.', { meType: 'Excel.Worksheet' });
		expect(got).toContain('Range');
		expect(got).toContain('Name');
	});

	it('resolves Me to the module host type (Workbook)', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = names(src, 'Me.', { meType: 'Excel.Workbook' });
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
	});

	it('merges Me document source members with its host surface', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Me.'), {
			meType: 'Excel.Workbook',
			meProjectType: 'ThisWorkbook',
			projectClassMembers: [
				{
					name: 'ThisWorkbook',
					kind: 'document',
					moduleName: 'ThisWorkbook',
					exhaustive: false,
					members: [
						{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' },
					],
				},
			],
		});
		expect(got.map((member) => member.name)).toContain('Hello');
		expect(got.map((member) => member.name)).toContain('AcceptAllChanges');
		expect(got.find((member) => member.name === 'Hello')?.surfaceExhaustive).toBe(true);
		expect(got.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive).toBe(true);
	});

	it('returns nothing for Me when module type is unknown', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		expect(names(src, 'Me.', {})).toEqual([]);
	});
});

describe('member completion - declared variables', () => {
	it('resolves a local Dim As Worksheet variable', () => {
		const src = 'Sub Test()\n    Dim ws As Worksheet\n    ws.\nEnd Sub\n';
		const got = names(src, 'ws.');
		expect(got).toContain('Range');
		expect(got).toContain('UsedRange');
	});

	it('uses the declared Worksheet type after assignment from Sheets(index)', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = Workbooks(1).Sheets(1)\n' +
			'    ws.\n' +
			'End Sub\n';
		const got = names(src, 'ws.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).not.toContain('ChartType');
	});

	it('uses the declared Chart type after assignment from Sheets(index)', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ch As Chart\n' +
			'    Set ch = Workbooks(1).Sheets(1)\n' +
			'    ch.\n' +
			'End Sub\n';
		const got = names(src, 'ch.');
		expect(got).toContain('ChartType');
		expect(got).toContain('SeriesCollection');
		expect(got).not.toContain('Range');
	});

	it('refines Object variables from simple Set assignments to known object expressions', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Workbooks(1).Worksheets(1)\n' +
			'    obj.\n' +
			'End Sub\n';
		const got = names(src, 'obj.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).not.toContain('ChartType');
	});

	it('keeps Object variables assigned from Sheets(index) on the merged item surface', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Workbooks(1).Sheets(1)\n' +
			'    obj.\n' +
			'End Sub\n';
		const got = names(src, 'obj.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).toContain('ChartType');
	});

	it('refines Variant variables from Set assignments and supports downstream chains', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Variant\n' +
			'    Set obj = Workbooks(1).Sheets(1)\n' +
			'    obj.Range("A1").\n' +
			'End Sub\n';
		const got = names(src, 'obj.Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('uses the latest preceding Set assignment for generic object variables', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Workbooks(1).Worksheets(1)\n' +
			'    Set obj = ActiveWorkbook\n' +
			'    obj.\n' +
			'End Sub\n';
		const got = names(src, 'obj.');
		expect(got).toContain('Save');
		expect(got).toContain('Worksheets');
		expect(got).not.toContain('Range');
	});

	it('resolves a parameter typed As Range', () => {
		const src = 'Sub Test(rng As Range)\n    rng.\nEnd Sub\n';
		const got = names(src, 'rng.');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('resolves a module-level variable', () => {
		const src = 'Private wb As Workbook\n\nSub Test()\n    wb.\nEnd Sub\n';
		const got = names(src, 'wb.');
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
	});

	it('resolves a variable declared inside a block', () => {
		const src =
			'Sub Test()\n    If True Then\n        Dim ws As Worksheet\n    End If\n    ws.\nEnd Sub\n';
		const got = names(src, 'ws.');
		expect(got).toContain('Range');
	});

	it('lets local declarations shadow host globals for member completion', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ActiveSheet As Person\n' +
			'    ActiveSheet.\n' +
			'End Sub\n';
		const got = names(src, 'ActiveSheet.', {
			projectClassMembers: [
				{
					name: 'Person',
					kind: 'class' as const,
					moduleName: 'Person',
					members: [
						{ name: 'Name', kind: 'property' as const, returns: 'String', moduleName: 'Person' },
						{ name: 'Save', kind: 'method' as const, moduleName: 'Person' },
					],
				},
			],
		});

		expect(got).toContain('Name');
		expect(got).toContain('Save');
		expect(got).not.toContain('Range');
	});

	it('lets untyped local declarations shadow host globals for member completion', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ActiveSheet\n' +
			'    ActiveSheet.\n' +
			'End Sub\n';

		expect(names(src, 'ActiveSheet.')).toEqual([]);
	});

	it('can refine untyped local host-global shadows from Set assignments for completion', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ActiveSheet\n' +
			'    Set ActiveSheet = ThisWorkbook\n' +
			'    ActiveSheet.\n' +
			'End Sub\n';
		const got = names(src, 'ActiveSheet.');

		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
		expect(got).not.toContain('Range');
	});
});

describe('member completion - workbook classes', () => {
	const projectClassMembers = [
		{
			name: 'Person',
			kind: 'class' as const,
			moduleName: 'Person',
			members: [
				{ name: 'Name', kind: 'property' as const, returns: 'String', moduleName: 'Person' },
				{
					name: 'Age',
					kind: 'property' as const,
					returns: 'Integer',
					writable: true,
					writeType: 'Integer',
					moduleName: 'Person',
					doc: {
						summary: 'Age in whole years.',
						params: [],
						source: 'inline' as const,
					},
				},
				{ name: 'Save', kind: 'method' as const, moduleName: 'Person' },
				{ name: 'Manager', kind: 'method' as const, returns: 'Person', moduleName: 'Person' },
			],
		},
	];

	it('offers members for a variable declared as a project class', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.\nEnd Sub\n';
		const got = names(src, 'p.', { projectClassMembers });
		expect(got).toContain('Name');
		expect(got).toContain('Save');
	});

	it('offers source-backed current class members through Me', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = names(src, 'Me.', { meProjectType: 'Person', projectClassMembers });
		expect(got).toContain('Name');
		expect(got).toContain('Save');
	});

	it('includes inline documentation for source-backed project class members', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.Ag\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.Ag'), {
			projectClassMembers,
		});
		const age = got.find((m) => m.name === 'Age');
		expect(age?.documentation).toContain('Age in whole years.');
		expect(age?.writable).toBe(true);
		expect(age?.writeType).toBe('Integer');
		expect(age?.surfaceExhaustive).toBe(true);
	});

	it('offers public class fields and excludes invalid public constants', () => {
		const person =
			'Public Age As Integer\n' +
			'Public Const Species As String = "Human"\n';
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		const src = 'Sub Test()\n    Dim p As Person\n    p.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.'), {
			projectClassMembers: index.projectClassMembers(),
		});
		const age = got.find((member) => member.name === 'Age');
		const species = got.find((member) => member.name === 'Species');
		expect(age?.writable).toBe(true);
		expect(age?.writeType).toBe('Integer');
		expect(species).toBeUndefined();
	});

	it('carries default-member attributes for source-backed project class members', () => {
		const person = [
			'Attribute VB_Name = "Person"',
			'Public Property Get Value() As String',
			'Attribute Value.VB_UserMemId = 0',
			'End Property',
		].join('\n');
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		const src = 'Sub Test()\n    Dim p As Person\n    p.Val\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.Val'), {
			projectClassMembers: index.projectClassMembers(),
		});
		const value = got.find((member) => member.name === 'Value');
		expect(value?.defaultMember).toBe(true);
		expect(value?.attributes?.[0]?.name).toBe('VB_UserMemId');
	});

	it('carries source definition locations for project class members', () => {
		const person =
			'Public Sub Save()\n' +
			'End Sub\n';
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		const src = 'Sub Test()\n    Dim p As Person\n    p.Sav\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.Sav'), {
			projectClassMembers: index.projectClassMembers(),
		});
		const save = got.find((member) => member.name === 'Save');
		expect(save?.definitions).toHaveLength(1);
		const def = save?.definitions?.[0];
		expect(def?.moduleName).toBe('Person');
		expect(def ? person.slice(def.nameSpan.start, def.nameSpan.end) : '').toBe('Save');
	});

	it('keeps same-named project class member definitions tied to the receiver type', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: 'Public Property Get FirstName() As String\nEnd Property\n',
		});
		index.setModule({
			moduleName: 'Class1',
			moduleKind: 'class',
			source: 'Public Property Get FirstName() As String\nEnd Property\n',
		});
		const src = [
			'Sub Test()',
			'    Dim p As Person',
			'    Dim c As Class1',
			'    p.First',
			'    c.First',
			'End Sub',
		].join('\n');
		const projectClassMembers = index.projectClassMembers();

		const personFirst = resolveMemberCompletions(src, dotOffset(src, 'p.First'), {
			projectClassMembers,
		}).find((member) => member.name === 'FirstName');
		const classFirst = resolveMemberCompletions(src, dotOffset(src, 'c.First'), {
			projectClassMembers,
		}).find((member) => member.name === 'FirstName');

		expect(personFirst?.definitions?.map((definition) => definition.moduleName)).toEqual(['Person']);
		expect(classFirst?.definitions?.map((definition) => definition.moduleName)).toEqual(['Class1']);
	});

	it('chains through project class members that return a project class', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.Manager.\nEnd Sub\n';
		const got = names(src, 'p.Manager.', { projectClassMembers });
		expect(got).toContain('Name');
		expect(got).toContain('Save');
	});

	it('does not resolve ambiguous project class member surfaces', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.\nEnd Sub\n';
		expect(names(src, 'p.', {
			projectClassMembers: [
				...projectClassMembers,
				{ ...projectClassMembers[0], moduleName: 'OtherPerson' },
			],
		})).toEqual([]);
	});

	it('marks source-only document module member surfaces as non-exhaustive', () => {
		const src = 'Sub Test()\n    Dim wb As ThisWorkbook\n    wb.H\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'wb.H'), {
			projectClassMembers: [
				{
					name: 'ThisWorkbook',
					kind: 'document',
					moduleName: 'ThisWorkbook',
					exhaustive: false,
					members: [
						{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' },
					],
				},
			],
		});
		expect(got.map((member) => member.name)).toContain('Hello');
		expect(got[0]?.surfaceExhaustive).toBe(false);
	});

	it('merges ThisWorkbook source members with the exhaustive Workbook host surface', () => {
		const src = 'Sub Test()\n    ThisWorkbook.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ThisWorkbook.'), {
			projectClassMembers: [
				{
					name: 'ThisWorkbook',
					kind: 'document',
					moduleName: 'ThisWorkbook',
					exhaustive: false,
					members: [
						{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' },
					],
				},
			],
		});
		expect(got.map((member) => member.name)).toContain('Hello');
		expect(got.map((member) => member.name)).toContain('AcceptAllChanges');
		expect(got.find((member) => member.name === 'Hello')?.surfaceExhaustive).toBe(true);
		expect(got.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive).toBe(true);
	});
});

describe('member definition resolution (references/rename path)', () => {
	const person = 'Public Sub Save()\nEnd Sub\n';

	function personCtx() {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		return { projectClassMembers: index.projectClassMembers() };
	}

	it('resolves definitions for a dotted occurrence without building completion rows', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.Save\nEnd Sub\n';
		const offset = src.indexOf('p.Save') + 'p.Save'.length;
		const got = resolveMemberDefinitionsAt(src, offset, 'Save', personCtx());
		expect(got).toHaveLength(1);
		expect(got[0].moduleName).toBe('Person');
		expect(person.slice(got[0].nameSpan.start, got[0].nameSpan.end)).toBe('Save');
	});

	it('bails on occurrences not preceded by a member-access dot', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    Save\nEnd Sub\n';
		const offset = src.indexOf('    Save') + '    Save'.length;
		expect(resolveMemberDefinitionsAt(src, offset, 'Save', personCtx())).toHaveLength(0);
	});

	it('follows a line continuation between the dot and the member name', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p. _\n        Save\nEnd Sub\n';
		const offset = src.indexOf('        Save') + '        Save'.length;
		const got = resolveMemberDefinitionsAt(src, offset, 'Save', personCtx());
		expect(got).toHaveLength(1);
	});

	it('accepts a precomputed prefix-token slice of the module', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.Save\nEnd Sub\n';
		const offset = src.indexOf('p.Save') + 'p.Save'.length;
		const moduleTokens = tokenize(src).filter((t) => t.kind !== 'comment');
		const prefixTokens = moduleTokens.filter((t) => t.end <= offset);
		const got = resolveMemberDefinitionsAt(src, offset, 'Save', personCtx(), prefixTokens);
		expect(got).toHaveLength(1);
		expect(got[0].moduleName).toBe('Person');
	});
});

describe('member completion - chaining', () => {
	it('walks a member chain through return types', () => {
		const src = 'Sub Test()\n    ThisWorkbook.ActiveSheet.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.ActiveSheet.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
	});

	it('walks chains across call parentheses', () => {
		const src = 'Sub Test(ws As Worksheet)\n    ws.Range("A1").Offset(1, 0).\nEnd Sub\n';
		const got = names(src, 'ws.Range("A1").Offset(1, 0).');
		expect(got).toContain('Value');
		expect(got).toContain('Resize');
	});

	it('resolves parenthesized receiver expressions', () => {
		const src = 'Sub Test(ws As Worksheet)\n    (ws.Range("A1")).Va\nEnd Sub\n';
		const got = names(src, ')).Va');
		expect(got).toContain('Value');
		expect(got).toContain('Value2');
	});

	it('walks through collection default Item for Worksheets(index).Range', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Worksheets(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.Worksheets(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('walks through global Workbooks(index).Worksheets(index).Range', () => {
		const src = 'Sub Test()\n    Workbooks(1).Worksheets(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'Workbooks(1).Worksheets(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Resize');
	});

	it('carries verified call signatures for callable member completions', () => {
		const src = 'Sub Test()\n    Workbooks(1).Worksheets(1).Ran\nEnd Sub\n';
		const got = resolveMemberCompletions(
			src,
			dotOffset(src, 'Workbooks(1).Worksheets(1).Ran'),
		);
		const range = got.find((m) => m.name === 'Range');
		expect(range?.signature).toBe('Range(Cell1, [Cell2]) As Range');
	});

	it('carries generated reference signatures and docs for promoted host members', () => {
		const src = 'Sub Test()\n    Application.Calc\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Application.Calc'));
		const calculate = got.find((m) => m.name === 'Calculate');
		expect(calculate?.signature).toBe('Calculate()');
		expect(calculate?.documentation).toContain('Calculates all open workbooks');
		expect(calculate?.doc?.source).toBe('external');
	});

	it('uses generated reference metadata ahead of fallback signatures', () => {
		const src = 'Sub Test()\n    Workbooks.Op\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Workbooks.Op'));
		const open = got.find((m) => m.name === 'Open');
		expect(open?.signature).toContain('[ReadOnly As Variant]');
		expect(open?.doc?.params.find((p) => p.name === 'ReadOnly')?.text).toContain(
			'read-only mode',
		);
	});

	it('walks through indexed Sheets into merged sheet object members', () => {
		const sheetSrc = 'Sub Test()\n    Workbooks(1).Sheets(1).\nEnd Sub\n';
		const sheetMembers = names(sheetSrc, 'Workbooks(1).Sheets(1).');
		expect(sheetMembers).toContain('Range');
		expect(sheetMembers).toContain('Cells');
		expect(sheetMembers).toContain('ChartType');

		const src = 'Sub Test()\n    Workbooks(1).Sheets(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'Workbooks(1).Sheets(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('walks through explicit Sheets.Item into worksheet members', () => {
		const src = 'Sub Test()\n    Workbooks(1).Sheets.Item(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'Workbooks(1).Sheets.Item(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Resize');
	});

	it('resolves Range.Worksheet back to a worksheet', () => {
		const src = 'Sub Test(rng As Range)\n    rng.Worksheet.\nEnd Sub\n';
		const got = names(src, 'rng.Worksheet.');
		expect(got).toContain('Cells');
		expect(got).toContain('Range');
	});
});

describe('member completion - user-defined types', () => {
	function userTypeIndex(): ProjectIndex {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: [
				'Public Type TPoint',
				'    X As Long',
				'    Y As Long',
				'End Type',
				'Public Type TBox',
				'    Corner As TPoint',
				'End Type',
				'Private Type THidden',
				'    Secret As String',
				'End Type',
			].join('\n'),
		});
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });
		return index;
	}

	it('offers fields for a variable declared as a visible UDT', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim p As TPoint\n    p.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.'), {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got.map((member) => member.name)).toEqual(['X', 'Y']);
		expect(got.find((member) => member.name === 'X')?.writeType).toBe('Long');
		expect(got.find((member) => member.name === 'X')?.surfaceExhaustive).toBe(true);
	});

	it('chains through nested UDT fields', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim box As TBox\n    box.Corner.\nEnd Sub\n';
		const got = names(src, 'box.Corner.', {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got).toContain('X');
		expect(got).toContain('Y');
	});

	it('resolves leading-dot fields inside With blocks', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim p As TPoint\n    With p\n        .\n    End With\nEnd Sub\n';
		const got = names(src, '        .', {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got).toContain('X');
		expect(got).toContain('Y');
	});

	it('does not expose private UDT fields outside their module', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim hidden As THidden\n    hidden.\nEnd Sub\n';
		const got = names(src, 'hidden.', {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got).toEqual([]);
	});
});

describe('member completion - With blocks', () => {
	it('resolves a leading-dot member against the active With receiver', () => {
		const src =
			'Sub Test(rng As Range)\n' +
			'    With rng\n' +
			'        .Va\n' +
			'    End With\n' +
			'End Sub\n';
		const got = names(src, '.Va');
		expect(got).toContain('Value');
		expect(got).toContain('Value2');
	});

	it('walks leading-dot chains from the active With receiver', () => {
		const src =
			'Sub Test(ws As Worksheet)\n' +
			'    With ws\n' +
			'        .Range("A1").\n' +
			'    End With\n' +
			'End Sub\n';
		const got = names(src, '.Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('resolves nested With receivers from outer leading-dot expressions', () => {
		const src =
			'Sub Test(ws As Worksheet)\n' +
			'    With ws\n' +
			'        With .Range("A1")\n' +
			'            .Va\n' +
			'        End With\n' +
			'    End With\n' +
			'End Sub\n';
		const got = names(src, '.Va');
		expect(got).toContain('Value');
		expect(got).toContain('Value2');
	});

	it('uses the enclosing procedure scope when resolving With receiver declarations', () => {
		const src =
			'Option Explicit\n' +
			'\n' +
			'Private ModuleValue As Long\n' +
			'\n' +
			'Sub Earlier()\n' +
			'End Sub\n' +
			'\n' +
			'Sub Test(ws As Worksheet)\n' +
			'    With ws\n' +
			'        .Ran\n' +
			'    End With\n' +
			'End Sub\n';
		const got = names(src, '.Ran');
		expect(got).toContain('Range');
	});

	it('keeps resolving inside an unfinished With block while editing', () => {
		const src =
			'Sub Test(rng As Range)\n' +
			'    With rng\n' +
			'        .Off\n' +
			'End Sub\n';
		const got = names(src, '.Off');
		expect(got).toContain('Offset');
	});
});

describe('member completion - negative cases', () => {
	it('returns nothing when not in a member-access position', () => {
		const src = 'Sub Test()\n    ThisWorkbook\nEnd Sub\n';
		const offset = src.indexOf('ThisWorkbook') + 'ThisWorkbook'.length;
		expect(resolveMemberCompletions(src, offset)).toEqual([]);
	});

	it('returns nothing for an unknown receiver', () => {
		const src = 'Sub Test()\n    Foo.\nEnd Sub\n';
		expect(names(src, 'Foo.')).toEqual([]);
	});

	it('returns nothing when chaining through a non-chainable member', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Name.\nEnd Sub\n';
		expect(names(src, 'ThisWorkbook.Name.')).toEqual([]);
	});

	it('host member lists are non-empty and well-formed', () => {
		for (const type of [
			'Excel.Application',
			'Excel.Workbook',
			'Excel.Worksheet',
			'Excel.Range',
			'Excel.Workbooks',
			'Excel.Worksheets',
			'Excel.Sheets',
		]) {
			const members = getHostMembers(type);
			expect(members.length).toBeGreaterThan(0);
			for (const mem of members) {
				expect(mem.name).toMatch(/^[A-Za-z_][A-Za-z0-9_]*$/);
				expect(['property', 'method']).toContain(mem.kind);
			}
		}
	});

	it('marks generated promoted Excel host member surfaces as exhaustive', () => {
		const src = 'Sub Test()\n    Application.Work\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Application.Work'));
		const workbooks = got.find((member) => member.name === 'Workbooks');
		expect(workbooks?.owner).toBe('Excel.Application');
		expect(workbooks?.surfaceExhaustive).toBe(true);
		expect(getHostType('Excel.Application')?.provenance).toContain('reference/excel/json/Application.json');
		for (const typeName of [
			'Excel.Workbooks',
			'Excel.Worksheets',
			'Excel.Sheets',
			'Excel.ListObject',
			'Excel.ListObjects',
			'Excel.ListRow',
			'Excel.ListRows',
			'Excel.ListColumn',
			'Excel.ListColumns',
			'Excel.Chart',
			'Excel.Charts',
			'Excel.ChartObject',
			'Excel.ChartObjects',
			'Excel.Shape',
			'Excel.Shapes',
			'Excel.Font',
			'Excel.Interior',
			'Excel.Border',
			'Excel.Borders',
			'Excel.Areas',
			'Excel.Hyperlink',
			'Excel.Hyperlinks',
			'Excel.FormatCondition',
			'Excel.FormatConditions',
			'Excel.Style',
			'Excel.Styles',
			'Excel.PageSetup',
			'Excel.Validation',
			'Excel.Name',
			'Excel.Names',
			'Excel.Window',
			'Excel.Windows',
		]) {
			expect(getHostType(typeName)?.exhaustive, typeName).toBe(true);
			expect(getHostType(typeName)?.provenance, typeName).toContain('reference/excel/json/');
		}
	});

	it('promotes non-exhaustive metadata surfaces without hard absence diagnostics', () => {
		for (const typeName of [
			'Excel.WorksheetFunction',
			'Excel.PivotTable',
			'Excel.PivotTables',
			'Excel.PivotField',
			'Excel.PivotFields',
			'Excel.PivotItem',
			'Excel.PivotItems',
			'Excel.PivotCache',
			'Excel.PivotCaches',
			'Excel.PivotFilter',
			'Excel.PivotFilters',
			'Excel.PivotCell',
			'Excel.PivotLayout',
			'Excel.PivotAxis',
			'Excel.PivotFormula',
			'Excel.PivotFormulas',
			'Excel.PivotLine',
			'Excel.PivotLineCells',
			'Excel.PivotLines',
			'Excel.PivotValueCell',
			'Excel.PivotTableChangeList',
			'Excel.PivotItemList',
			'Excel.CalculatedFields',
			'Excel.CalculatedItems',
			'Excel.CubeField',
			'Excel.CubeFields',
			'Excel.QueryTable',
			'Excel.QueryTables',
			'Excel.Parameter',
			'Excel.Parameters',
			'Excel.SeriesCollection',
			'Excel.FullSeriesCollection',
			'Excel.Series',
			'Excel.Axes',
			'Excel.Axis',
			'Excel.AxisTitle',
			'Excel.ChartFormat',
			'Excel.TickLabels',
			'Excel.Gridlines',
			'Excel.DisplayUnitLabel',
			'Excel.LeaderLines',
			'Excel.ErrorBars',
			'Excel.Points',
			'Excel.Point',
			'Excel.Trendlines',
			'Excel.Trendline',
			'Excel.DataLabels',
			'Excel.DataLabel',
			'Excel.ShapeRange',
			'Excel.GroupShapes',
			'Excel.Comment',
			'Excel.Comments',
			'Excel.CommentThreaded',
			'Excel.CommentsThreaded',
			'Excel.Sort',
			'Excel.SortFields',
			'Excel.SortField',
			'Excel.AutoFilter',
			'Excel.Filters',
			'Excel.Filter',
			'Excel.Icon',
			'Excel.IconSet',
			'Excel.IconSets',
			'Excel.OLEObject',
			'Excel.OLEObjects',
			'Excel.Button',
			'Excel.Buttons',
			'Excel.CheckBox',
			'Excel.CheckBoxes',
			'Excel.DropDown',
			'Excel.DropDowns',
			'Excel.OptionButton',
			'Excel.OptionButtons',
			'Excel.GroupObject',
			'Excel.GroupObjects',
			'Excel.GroupBox',
			'Excel.GroupBoxes',
			'Excel.Label',
			'Excel.Labels',
			'Excel.ListBox',
			'Excel.ListBoxes',
			'Excel.ScrollBar',
			'Excel.ScrollBars',
			'Excel.Spinner',
			'Excel.Spinners',
			'Excel.EditBox',
			'Excel.EditBoxes',
			'Excel.TextBox',
			'Excel.TextBoxes',
			'Excel.WorkbookConnection',
			'Excel.Connections',
			'Excel.OLEDBConnection',
			'Excel.ODBCConnection',
			'Excel.Ranges',
			'Excel.ModelConnection',
			'Excel.WorksheetDataConnection',
			'Excel.TextConnection',
			'Excel.DataFeedConnection',
			'Excel.ModelTables',
			'Excel.ModelTable',
			'Excel.CalculatedMember',
			'Excel.CalculatedMembers',
			'Excel.ModelTableColumn',
			'Excel.ModelTableColumns',
			'Excel.SlicerCache',
			'Excel.SlicerCaches',
			'Excel.Slicer',
			'Excel.Slicers',
			'Excel.SlicerItem',
			'Excel.SlicerItems',
			'Excel.SlicerCacheLevel',
			'Excel.SlicerCacheLevels',
			'Excel.SlicerPivotTables',
			'Excel.TimelineState',
			'Excel.TimelineViewState',
			'Excel.FillFormat',
			'Excel.LineFormat',
			'Excel.TextFrame',
			'Excel.TextFrame2',
			'Excel.PictureFormat',
			'Excel.ShadowFormat',
			'Excel.ThreeDFormat',
			'Excel.ConnectorFormat',
			'Excel.CalloutFormat',
			'Excel.ShapeNodes',
			'Excel.ShapeNode',
			'Excel.ColorFormat',
			'Excel.ChartTitle',
			'Excel.ChartArea',
			'Excel.PlotArea',
			'Excel.Legend',
			'Excel.LegendEntry',
			'Excel.LegendEntries',
			'Excel.LegendKey',
			'Excel.ChartGroup',
			'Excel.ChartGroups',
			'Excel.DataTable',
			'Excel.Walls',
			'Excel.Floor',
			'Excel.SeriesLines',
			'Excel.DownBars',
			'Excel.DropLines',
			'Excel.HiLoLines',
			'Excel.UpBars',
			'Excel.Databar',
			'Excel.ColorScale',
			'Excel.ColorScaleCriteria',
			'Excel.ColorScaleCriterion',
			'Excel.IconSetCondition',
			'Excel.AboveAverage',
			'Excel.Top10',
			'Excel.UniqueValues',
			'Excel.CellFormat',
			'Excel.DisplayFormat',
			'Excel.FormatColor',
			'Excel.ConditionValue',
			'Excel.DataBarBorder',
			'Excel.NegativeBarFormat',
			'Excel.IconCriteria',
			'Excel.IconCriterion',
			'Excel.DrawingObjects',
			'Excel.Drawing',
			'Excel.Drawings',
			'Excel.Picture',
			'Excel.Pictures',
			'Excel.Line',
			'Excel.Lines',
			'Excel.Rectangle',
			'Excel.Rectangles',
			'Excel.Oval',
			'Excel.Ovals',
			'Excel.Arc',
			'Excel.Arcs',
			'Excel.SparklineGroups',
			'Excel.SparklineGroup',
			'Excel.Sparkline',
			'Excel.SparkPoints',
			'Excel.SparkColor',
			'Excel.SparkAxes',
			'Excel.SparkHorizontalAxis',
			'Excel.SparkVerticalAxis',
			'Excel.XmlMap',
			'Excel.XmlMaps',
			'Excel.XPath',
			'Excel.PublishObject',
			'Excel.PublishObjects',
			'Excel.WebOptions',
			'Excel.DefaultWebOptions',
			'Excel.XmlNamespace',
			'Excel.XmlNamespaces',
			'Excel.XmlSchema',
			'Excel.XmlSchemas',
			'Excel.XmlDataBinding',
		]) {
			expect(getHostType(typeName)?.provenance, typeName).toContain('reference/excel/json/');
			expect(getHostType(typeName)?.exhaustive, typeName).not.toBe(true);
		}
	});

	it('marks generated Worksheet host member surfaces as exhaustive', () => {
		const src = 'Sub Test()\n    ActiveSheet.Named\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ActiveSheet.Named'));
		const namedSheetViews = got.find((member) => member.name === 'NamedSheetViews');
		expect(namedSheetViews?.owner).toBe('Excel.Worksheet');
		expect(namedSheetViews?.surfaceExhaustive).toBe(true);
	});

	it('marks generated Range host member surfaces as exhaustive', () => {
		const src = 'Sub Test()\n    ActiveCell.Val\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ActiveCell.Val'));
		const value2 = got.find((member) => member.name === 'Value2');
		expect(value2?.owner).toBe('Excel.Range');
		expect(value2?.surfaceExhaustive).toBe(true);
	});

	it('marks dump-backed Workbook host member surfaces as exhaustive', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Sav\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ThisWorkbook.Sav'));
		const save = got.find((member) => member.name === 'Save');
		expect(save?.owner).toBe('Excel.Workbook');
		expect(save?.surfaceExhaustive).toBe(true);
	});

	it('can mark a verified exhaustive host member surface', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					exhaustive: true,
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src = 'Sub Test()\n    Thing.K\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Thing.K'), { model });
		expect(got[0]?.name).toBe('Known');
		expect(got[0]?.surfaceExhaustive).toBe(true);
	});
});
