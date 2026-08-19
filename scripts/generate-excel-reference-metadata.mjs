import fs from 'node:fs';
import path from 'node:path';
import {
	cleanText,
	collectConstants,
	readReferenceDumps,
	renderConstant,
} from './reference-generator-utils.mjs';

const root = process.cwd();
const jsonDir = path.join(root, 'reference', 'excel', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'excelReferenceMembers.ts');
const coveragePath = path.join(root, 'docs', 'excel_reference_coverage.md');

// Keep runtime promotion explicit. The generator scans the full Excel corpus for
// coverage, but only promotedTypes become checked-in host metadata. Only
// hardDiagnosticTypes are allowed to prove hard member-not-found absence.
const hardDiagnosticTypes = [
	'Application',
	'Workbook',
	'Worksheet',
	'Range',
	'Workbooks',
	'Worksheets',
	'Sheets',
	'Window',
	'Windows',
	'Name',
	'Names',
	'ListObject',
	'ListObjects',
	'ListRow',
	'ListRows',
	'ListColumn',
	'ListColumns',
	'Chart',
	'Charts',
	'ChartObject',
	'ChartObjects',
	'Shape',
	'Shapes',
	'Font',
	'Interior',
	'Border',
	'Borders',
	'Areas',
	'Hyperlink',
	'Hyperlinks',
	'FormatCondition',
	'FormatConditions',
	'Style',
	'Styles',
	'PageSetup',
	'Validation',
];

const promotedTypes = [
	...hardDiagnosticTypes,
	// The hidden Global interface: the members VBA calls bare (Union,
	// Intersect, Evaluate). Resolved by resolveHostGlobalMember (issue #34).
	'Global',
	'WorksheetFunction',
	'PivotTable',
	'PivotTables',
	'PivotField',
	'PivotFields',
	'PivotItem',
	'PivotItems',
	'PivotCache',
	'PivotCaches',
	'PivotFilter',
	'PivotFilters',
	'PivotCell',
	'PivotLayout',
	'PivotAxis',
	'PivotFormula',
	'PivotFormulas',
	'PivotLine',
	'PivotLineCells',
	'PivotLines',
	'PivotValueCell',
	'PivotTableChangeList',
	'PivotItemList',
	'CalculatedFields',
	'CalculatedItems',
	'CubeField',
	'CubeFields',
	'QueryTable',
	'QueryTables',
	'Parameter',
	'Parameters',
	'SeriesCollection',
	'FullSeriesCollection',
	'Series',
	'Axes',
	'Axis',
	'AxisTitle',
	'ChartFormat',
	'TickLabels',
	'Gridlines',
	'DisplayUnitLabel',
	'LeaderLines',
	'ErrorBars',
	'Points',
	'Point',
	'Trendlines',
	'Trendline',
	'DataLabels',
	'DataLabel',
	'ShapeRange',
	'GroupShapes',
	'Comment',
	'Comments',
	'CommentThreaded',
	'CommentsThreaded',
	'Sort',
	'SortFields',
	'SortField',
	'AutoFilter',
	'Filters',
	'Filter',
	'Icon',
	'IconSet',
	'IconSets',
	'OLEObject',
	'OLEObjects',
	'Button',
	'Buttons',
	'CheckBox',
	'CheckBoxes',
	'DropDown',
	'DropDowns',
	'OptionButton',
	'OptionButtons',
	'GroupObject',
	'GroupObjects',
	'GroupBox',
	'GroupBoxes',
	'Label',
	'Labels',
	'ListBox',
	'ListBoxes',
	'ScrollBar',
	'ScrollBars',
	'Spinner',
	'Spinners',
	'EditBox',
	'EditBoxes',
	'TextBox',
	'TextBoxes',
	'WorkbookConnection',
	'Connections',
	'OLEDBConnection',
	'ODBCConnection',
	'Ranges',
	'ModelConnection',
	'WorksheetDataConnection',
	'TextConnection',
	'DataFeedConnection',
	'ModelTables',
	'ModelTable',
	'CalculatedMember',
	'CalculatedMembers',
	'ModelTableColumn',
	'ModelTableColumns',
	'SlicerCache',
	'SlicerCaches',
	'Slicer',
	'Slicers',
	'SlicerItem',
	'SlicerItems',
	'SlicerCacheLevel',
	'SlicerCacheLevels',
	'SlicerPivotTables',
	'TimelineState',
	'TimelineViewState',
	'FillFormat',
	'LineFormat',
	'TextFrame',
	'TextFrame2',
	'PictureFormat',
	'ShadowFormat',
	'ThreeDFormat',
	'ConnectorFormat',
	'CalloutFormat',
	'ShapeNodes',
	'ShapeNode',
	'ColorFormat',
	'ChartTitle',
	'ChartArea',
	'PlotArea',
	'Legend',
	'LegendEntry',
	'LegendEntries',
	'LegendKey',
	'ChartGroup',
	'ChartGroups',
	'DataTable',
	'Walls',
	'Floor',
	'SeriesLines',
	'DownBars',
	'DropLines',
	'HiLoLines',
	'UpBars',
	'Databar',
	'ColorScale',
	'ColorScaleCriteria',
	'ColorScaleCriterion',
	'IconSetCondition',
	'AboveAverage',
	'Top10',
	'UniqueValues',
	'CellFormat',
	'DisplayFormat',
	'FormatColor',
	'ConditionValue',
	'DataBarBorder',
	'NegativeBarFormat',
	'IconCriteria',
	'IconCriterion',
	'DrawingObjects',
	'Drawing',
	'Drawings',
	'Picture',
	'Pictures',
	'Line',
	'Lines',
	'Rectangle',
	'Rectangles',
	'Oval',
	'Ovals',
	'Arc',
	'Arcs',
	'SparklineGroups',
	'SparklineGroup',
	'Sparkline',
	'SparkPoints',
	'SparkColor',
	'SparkAxes',
	'SparkHorizontalAxis',
	'SparkVerticalAxis',
	'XmlMap',
	'XmlMaps',
	'XPath',
	'PublishObject',
	'PublishObjects',
	'WebOptions',
	'DefaultWebOptions',
	'XmlNamespace',
	'XmlNamespaces',
	'XmlSchema',
	'XmlSchemas',
	'XmlDataBinding',
];

const primitiveTypes = new Set([
	'Boolean',
	'Byte',
	'Currency',
	'Date',
	'Decimal',
	'Double',
	'Integer',
	'Long',
	'LongLong',
	'LongPtr',
	'Object',
	'Single',
	'String',
	'Variant',
	'void',
	'IUnknown',
]);

const dumps = readReferenceDumps(jsonDir);

function normalizeTypeName(typeName) {
	if (!typeName || typeof typeName !== 'string') {
		return undefined;
	}
	return typeName.replace(/\(.*\)$/g, '').trim();
}

function excelQualifiedReturn(typeName) {
	const clean = normalizeTypeName(typeName);
	if (!clean || primitiveTypes.has(clean)) {
		return undefined;
	}
	if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(clean)) {
		return undefined;
	}
	if (!dumps.has(clean)) {
		return undefined;
	}
	return `Excel.${clean}`;
}

function memberReturn(raw, kind) {
	return kind === 'method' ? raw.returns : raw.type ?? raw.returns;
}

function memberDoc(raw) {
	const summary = cleanText(raw.description);
	const remarks = cleanText(raw.remarks);
	const params = (raw.parameters ?? [])
		.map((param) => {
			if (!param) {
				return undefined;
			}
			const text = cleanText(param.description);
			if (!param?.name || !text) {
				return undefined;
			}
			const type = cleanText(param.type);
			return {
				name: param.name,
				text,
				...(type ? { type } : {}),
			};
		})
		.filter(Boolean);
	if (!summary && !remarks && params.length === 0) {
		return undefined;
	}
	return {
		...(summary ? { summary } : {}),
		params,
		...(remarks ? { remarks } : {}),
		source: 'external',
	};
}

function memberFrom(raw, kind) {
	const qualifiedReturn = excelQualifiedReturn(memberReturn(raw, kind));
	const signature = cleanText(raw.signature);
	const doc = memberDoc(raw);
	return {
		name: raw.name,
		kind,
		...(qualifiedReturn ? { returns: qualifiedReturn } : {}),
		...(signature ? { signature } : {}),
		...(doc ? { doc } : {}),
	};
}

function collectMembers(typeDump, options = {}) {
	const includeEvents = options.includeEvents === true;
	const byName = new Map();
	const duplicateNames = new Set();
	const add = (raw, kind) => {
		if (!raw?.name) {
			return;
		}
		const key = raw.name.toLowerCase();
		if (byName.has(key)) {
			duplicateNames.add(raw.name);
			return;
		}
		byName.set(key, memberFrom(raw, kind));
	};

	for (const item of typeDump.properties ?? []) {
		add(item, 'property');
	}
	for (const item of typeDump.methods ?? []) {
		add(item, 'method');
	}
	if (includeEvents) {
		for (const item of typeDump.events ?? []) {
			add(item, 'event');
		}
	}
	return {
		members: [...byName.values()].sort((a, b) => a.name.localeCompare(b.name, 'en')),
		duplicateNames: [...duplicateNames].sort((a, b) => a.localeCompare(b, 'en')),
	};
}

function renderMember(member) {
	const parts = [
		`name: ${JSON.stringify(member.name)}`,
		`kind: ${JSON.stringify(member.kind)}`,
	];
	if (member.returns) {
		parts.push(`returns: ${JSON.stringify(member.returns)}`);
	}
	if (member.signature) {
		parts.push(`signature: ${JSON.stringify(member.signature)}`);
	}
	if (member.doc) {
		parts.push(`doc: ${JSON.stringify(member.doc)}`);
	}
	return `\t\t{ ${parts.join(', ')} },`;
}

function provenanceFor(typeName) {
	const entry = dumps.get(typeName);
	if (!entry) {
		throw new Error(`Missing Excel reference dump for promoted type ${typeName}`);
	}
	const { dump, fileName } = entry;
	return `${dump.library}; ${dump.guid}; reference/excel/json/${fileName}`;
}

function renderPromotedMemberSet(typeName) {
	const entry = dumps.get(typeName);
	if (!entry) {
		throw new Error(`Missing Excel reference dump for promoted type ${typeName}`);
	}
	const { members } = collectMembers(entry.dump);
	return `\t${JSON.stringify(typeName)}: [\n${members.map(renderMember).join('\n')}\n\t],`;
}

function typeCoverage(entry) {
	const { dump, fileName } = entry;
	const properties = dump.properties?.length ?? 0;
	const methods = dump.methods?.length ?? 0;
	const events = dump.events?.length ?? 0;
	const constants = dump.constants?.length ?? 0;
	const totalMembers = properties + methods + events;
	const { members, duplicateNames } = collectMembers(dump, { includeEvents: true });
	const objectSurfaceMembers = collectMembers(dump).members.length;
	const memberRows = [
		...(dump.properties ?? []).map((item) => ({ item, kind: 'property' })),
		...(dump.methods ?? []).map((item) => ({ item, kind: 'method' })),
		...(dump.events ?? []).map((item) => ({ item, kind: 'event' })),
	];
	const membersWithReturnType = memberRows.filter(({ item, kind }) =>
		memberReturn(item, kind),
	).length;
	const membersWithQualifiedReturnType = members.filter((member) => member.returns).length;
	const membersWithSignature = memberRows.filter(({ item }) => item.signature).length;
	const parameters = memberRows.reduce(
		(sum, { item }) => sum + (item.parameters?.length ?? 0),
		0,
	);
	const typedParameters = memberRows.reduce(
		(sum, { item }) =>
			sum + (item.parameters ?? []).filter((param) => Boolean(param.type)).length,
		0,
	);
	return {
		name: dump.name,
		kind: dump.kind ?? 'Unknown',
		fileName,
		properties,
		methods,
		events,
		constants,
		totalMembers,
		uniqueMembers: members.length,
		objectSurfaceMembers,
		duplicateNames: duplicateNames.length,
		membersWithReturnType,
		membersWithQualifiedReturnType,
		membersWithSignature,
		parameters,
		typedParameters,
	};
}

const coverage = [...dumps.values()]
	.map(typeCoverage)
	.sort((a, b) => a.name.localeCompare(b.name, 'en'));
const constants = collectConstants(dumps);

function countWhere(predicate) {
	return coverage.filter(predicate).length;
}

function sumOf(field, rows = coverage) {
	return rows.reduce((sum, row) => sum + row[field], 0);
}

function markdownTable(headers, rows) {
	return [
		`| ${headers.join(' | ')} |`,
		`| ${headers.map(() => '---').join(' | ')} |`,
		...rows.map((row) => `| ${row.join(' | ')} |`),
	].join('\n');
}

function renderCoverageMarkdown() {
	const objectRows = coverage.filter((row) => row.totalMembers > 0);
	const enumRows = coverage.filter((row) => row.constants > 0);
	const promotedRows = promotedTypes.map((typeName) => {
		const row = coverage.find((item) => item.name === typeName);
		if (!row) {
			throw new Error(`Missing coverage row for promoted type ${typeName}`);
		}
		return row;
	});
	const largestRows = [...objectRows]
		.sort(
			(a, b) =>
				b.uniqueMembers - a.uniqueMembers ||
				a.name.localeCompare(b.name, 'en'),
		)
		.slice(0, 30);

	return `# Excel Reference Coverage

Generated by \`npm run generate:reference:excel\` from \`reference/excel/json\`.

## Summary

${markdownTable(
	['Metric', 'Count'],
	[
		['JSON files scanned', coverage.length],
		['Object-like type dumps', countWhere((row) => row.totalMembers > 0)],
		['Enumeration dumps', enumRows.length],
		['Promoted runtime types', promotedTypes.length],
		['Hard diagnostic runtime types', hardDiagnosticTypes.length],
		['Raw properties', sumOf('properties')],
		['Raw methods', sumOf('methods')],
		['Raw events', sumOf('events')],
		['Raw enum constants', sumOf('constants')],
		['Raw member rows', sumOf('totalMembers')],
		['Unique object member names', sumOf('uniqueMembers')],
		['Runtime object-surface members', sumOf('objectSurfaceMembers')],
		['Member rows with signatures', sumOf('membersWithSignature')],
		['Member rows with return/type data', sumOf('membersWithReturnType')],
		['Parameters with type data', sumOf('typedParameters')],
	],
)}

## Promoted Runtime Types

${markdownTable(
	[
		'Type',
		'Members',
		'Runtime object members',
		'Properties',
		'Methods',
		'Events',
		'Return/type rows',
		'Qualified returns',
		'Signatures',
		'Duplicate names',
	],
	promotedRows.map((row) => [
		row.name,
		row.uniqueMembers,
		row.objectSurfaceMembers,
		row.properties,
		row.methods,
		row.events,
		row.membersWithReturnType,
		row.membersWithQualifiedReturnType,
		row.membersWithSignature,
		row.duplicateNames,
	]),
)}

## Largest Object Surfaces

${markdownTable(
	[
		'Type',
		'Kind',
		'Members',
		'Runtime object members',
		'Properties',
		'Methods',
		'Events',
		'Signatures',
	],
	largestRows.map((row) => [
		row.name,
		row.kind,
		row.uniqueMembers,
		row.objectSurfaceMembers,
		row.properties,
		row.methods,
		row.events,
		row.membersWithSignature,
	]),
)}

## Notes

- Runtime extension code does not read \`reference/\`; promoted metadata is checked in under \`src/\`.
- Promoted members and enum constants preserve available signatures, documentation summaries/parameter notes, types, and values for language-service surfaces and diagnostics.
- Excel events are counted for coverage but are intentionally not emitted into object-member surfaces; VBE does not expose events as callable object methods/properties. Event handler authoring uses a separate module-scoped metadata path.
- Completion may use partial metadata, but hard \`member-not-found\` diagnostics require a promoted exhaustive surface.
- Promotion remains type-by-type so each host surface can get representative tests and oracle controls before red diagnostics rely on absence.
- Hard \`member-not-found\` diagnostics are currently limited to these promoted
  runtime types: ${hardDiagnosticTypes.map((typeName) => `\`${typeName}\``).join(', ')}.
- Promoted metadata types outside the hard-diagnostic set remain completion,
  hover, signature, and receiver-chain surfaces until oracle evidence shows
  unknown members are compile-rejected.
`;
}

function renderOutput() {
	const provenanceEntries = promotedTypes
		.map((typeName) => `\t${JSON.stringify(typeName)}: ${JSON.stringify(provenanceFor(typeName))},`)
		.join('\n');
	const memberSetEntries = promotedTypes.map(renderPromotedMemberSet).join('\n');
	const constantEntries = constants.map(renderConstant).join('\n');
	const workbookProvenance = provenanceFor('Workbook');

	return `// Generated from reference/excel/json. Do not hand-edit member names here.
// Regenerate from the repo-local reference dump with \`npm run generate:reference:excel\`.

import type { HostConstant, HostMember } from './excelObjectModel';

export const EXCEL_REFERENCE_PROMOTED_TYPES = ${JSON.stringify(promotedTypes)} as const;

export const EXCEL_REFERENCE_HARD_DIAGNOSTIC_TYPES = ${JSON.stringify(hardDiagnosticTypes)} as const;

export const EXCEL_REFERENCE_PROVENANCE: Record<string, string> = {
${provenanceEntries}
};

export const EXCEL_REFERENCE_MEMBER_SETS: Record<string, readonly HostMember[]> = {
${memberSetEntries}
};

export const EXCEL_REFERENCE_ENUM_CONSTANTS: Record<string, HostConstant> = {
${constantEntries}
};

export const EXCEL_WORKBOOK_REFERENCE_PROVENANCE = ${JSON.stringify(workbookProvenance)};

export const EXCEL_WORKBOOK_REFERENCE_MEMBERS = EXCEL_REFERENCE_MEMBER_SETS.Workbook;
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, renderOutput(), 'utf8');
fs.mkdirSync(path.dirname(coveragePath), { recursive: true });
fs.writeFileSync(coveragePath, renderCoverageMarkdown(), 'utf8');

console.log(
	`Generated ${promotedTypes.length} promoted Excel reference type(s) at ${outputPath}`,
);
console.log(`Wrote Excel reference coverage report at ${coveragePath}`);
