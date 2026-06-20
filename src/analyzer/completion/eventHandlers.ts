// Module-scoped Excel event-handler completions.
//
// Excel events are not callable object members in VBA; `ThisWorkbook.Open()`
// and `ActiveSheet.Change()` are invalid member calls. The editor still needs
// a first-class authoring path for the procedure declarations that VBE wires up
// by exact module/object prefix and signature.

import { parseModule } from '../parser/parseModule';
import type { ProcedureNode } from '../parser/nodes';
import type { ModuleSymbolKind } from '../symbols/symbolModel';

export type EventHandlerDocumentType = 'workbook' | 'worksheet' | 'chart';

export interface EventHandlerCompletionContext {
	moduleName?: string;
	moduleKind?: ModuleSymbolKind;
	documentType?: EventHandlerDocumentType;
}

export interface EventHandlerCompletion {
	name: string;
	signature: string;
	detail: string;
	documentation: string;
	insertText: string;
}

interface EventHandlerDefinition {
	name: string;
	params: string;
	owner: 'Workbook' | 'Worksheet' | 'Chart';
	description: string;
}

export interface EventHandlerProcedureMatch {
	name: string;
	signature: string;
	owner: EventHandlerDefinition['owner'];
	documentType: EventHandlerDocumentType;
	description: string;
}

interface LineCompletionContext {
	currentWord: string;
	insertMode: 'fullProcedure' | 'declarationTail';
}

const WORKBOOK_EVENTS: readonly EventHandlerDefinition[] = [
	{
		name: 'Workbook_Open',
		params: '',
		owner: 'Workbook',
		description: 'Occurs when the workbook is opened.',
	},
	{
		name: 'Workbook_BeforeClose',
		params: 'Cancel As Boolean',
		owner: 'Workbook',
		description: 'Occurs before the workbook closes.',
	},
	{
		name: 'Workbook_BeforeSave',
		params: 'ByVal SaveAsUI As Boolean, Cancel As Boolean',
		owner: 'Workbook',
		description: 'Occurs before the workbook is saved.',
	},
	{
		name: 'Workbook_BeforePrint',
		params: 'Cancel As Boolean',
		owner: 'Workbook',
		description: 'Occurs before the workbook is printed.',
	},
	{
		name: 'Workbook_Activate',
		params: '',
		owner: 'Workbook',
		description: 'Occurs when the workbook is activated.',
	},
	{
		name: 'Workbook_Deactivate',
		params: '',
		owner: 'Workbook',
		description: 'Occurs when the workbook is deactivated.',
	},
	{
		name: 'Workbook_NewSheet',
		params: 'ByVal Sh As Object',
		owner: 'Workbook',
		description: 'Occurs when a new sheet is created in the workbook.',
	},
	{
		name: 'Workbook_SheetChange',
		params: 'ByVal Sh As Object, ByVal Target As Range',
		owner: 'Workbook',
		description: 'Occurs when cells are changed on any worksheet in the workbook.',
	},
	{
		name: 'Workbook_SheetCalculate',
		params: 'ByVal Sh As Object',
		owner: 'Workbook',
		description: 'Occurs after any worksheet is recalculated or chart data is plotted.',
	},
	{
		name: 'Workbook_SheetSelectionChange',
		params: 'ByVal Sh As Object, ByVal Target As Range',
		owner: 'Workbook',
		description: 'Occurs when the selection changes on any worksheet in the workbook.',
	},
];

const WORKSHEET_EVENTS: readonly EventHandlerDefinition[] = [
	{
		name: 'Worksheet_Change',
		params: 'ByVal Target As Range',
		owner: 'Worksheet',
		description: 'Occurs when cells on the worksheet are changed.',
	},
	{
		name: 'Worksheet_SelectionChange',
		params: 'ByVal Target As Range',
		owner: 'Worksheet',
		description: 'Occurs when the selection changes on the worksheet.',
	},
	{
		name: 'Worksheet_Calculate',
		params: '',
		owner: 'Worksheet',
		description: 'Occurs after the worksheet is recalculated.',
	},
	{
		name: 'Worksheet_Activate',
		params: '',
		owner: 'Worksheet',
		description: 'Occurs when the worksheet is activated.',
	},
	{
		name: 'Worksheet_Deactivate',
		params: '',
		owner: 'Worksheet',
		description: 'Occurs when the worksheet is deactivated.',
	},
	{
		name: 'Worksheet_BeforeDoubleClick',
		params: 'ByVal Target As Range, Cancel As Boolean',
		owner: 'Worksheet',
		description: 'Occurs before the default double-click action on the worksheet.',
	},
	{
		name: 'Worksheet_BeforeRightClick',
		params: 'ByVal Target As Range, Cancel As Boolean',
		owner: 'Worksheet',
		description: 'Occurs before the default right-click action on the worksheet.',
	},
	{
		name: 'Worksheet_FollowHyperlink',
		params: 'ByVal Target As Hyperlink',
		owner: 'Worksheet',
		description: 'Occurs when a hyperlink on the worksheet is selected.',
	},
];

const CHART_EVENTS: readonly EventHandlerDefinition[] = [
	{
		name: 'Chart_Activate',
		params: '',
		owner: 'Chart',
		description: 'Occurs when the chart is activated.',
	},
	{
		name: 'Chart_Deactivate',
		params: '',
		owner: 'Chart',
		description: 'Occurs when the chart is deactivated.',
	},
	{
		name: 'Chart_Calculate',
		params: '',
		owner: 'Chart',
		description: 'Occurs after the chart plots new or changed data.',
	},
	{
		name: 'Chart_Resize',
		params: '',
		owner: 'Chart',
		description: 'Occurs when the chart is resized.',
	},
	{
		name: 'Chart_Select',
		params: 'ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long',
		owner: 'Chart',
		description: 'Occurs when a chart element is selected.',
	},
	{
		name: 'Chart_SeriesChange',
		params: 'ByVal SeriesIndex As Long, ByVal PointIndex As Long',
		owner: 'Chart',
		description: 'Occurs when the value of a chart data point is changed.',
	},
	{
		name: 'Chart_BeforeDoubleClick',
		params: 'ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean',
		owner: 'Chart',
		description: 'Occurs before the default double-click action on the chart.',
	},
	{
		name: 'Chart_BeforeRightClick',
		params: 'Cancel As Boolean',
		owner: 'Chart',
		description: 'Occurs before the default right-click action on the chart.',
	},
	{
		name: 'Chart_MouseDown',
		params: 'ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long',
		owner: 'Chart',
		description: 'Occurs when a mouse button is pressed over the chart.',
	},
	{
		name: 'Chart_MouseMove',
		params: 'ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long',
		owner: 'Chart',
		description: 'Occurs when the mouse pointer moves over the chart.',
	},
	{
		name: 'Chart_MouseUp',
		params: 'ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long',
		owner: 'Chart',
		description: 'Occurs when a mouse button is released over the chart.',
	},
];

/**
 * Resolves event-procedure stub completions at module level.
 *
 * Results are intentionally scoped by the module being edited. Standard/class
 * modules can still define same-named procedures, but those are ordinary
 * procedures and are not offered here as wired event handlers.
 */
export function resolveEventHandlerCompletions(
	source: string,
	offset: number,
	ctx: EventHandlerCompletionContext = {},
): EventHandlerCompletion[] {
	if (ctx.moduleKind !== 'document') {
		return [];
	}

	const line = lineCompletionContext(source, offset);
	if (!line) {
		return [];
	}

	const documentType = eventHandlerDocumentTypeForContext(ctx);
	const definitions = definitionsForDocumentType(documentType);
	if (definitions.length === 0) {
		return [];
	}

	const parsed = parseModule(source);
	const procedures = parsed.members.filter((m): m is ProcedureNode => m.kind === 'Procedure');
	if (insideClosedProcedure(procedures, offset)) {
		return [];
	}

	const existing = new Set(
		procedures
			.map((proc) => proc.name.toLowerCase())
			.filter((name) => name.length > 0),
	);
	const prefix = line.currentWord.toLowerCase();

	return definitions
		.filter((def) => !existing.has(def.name.toLowerCase()))
		.filter((def) => !prefix || def.name.toLowerCase().startsWith(prefix))
		.map((def) => toCompletion(def, line.insertMode));
}

export function eventHandlerDocumentTypeForContext(
	ctx: EventHandlerCompletionContext,
): EventHandlerDocumentType | undefined {
	if (ctx.moduleKind !== 'document') {
		return undefined;
	}
	return ctx.documentType ?? inferDocumentType(ctx.moduleName);
}

export function eventHandlerProcedureForName(
	name: string,
): EventHandlerProcedureMatch | undefined {
	const definition = ALL_EVENT_DEFINITIONS.find(
		(def) => def.name.toLowerCase() === name.toLowerCase(),
	);
	if (!definition) {
		return undefined;
	}
	const documentType = documentTypeForOwner(definition.owner);
	if (!documentType) {
		return undefined;
	}
	return {
		name: definition.name,
		signature: `${definition.name}(${definition.params})`,
		owner: definition.owner,
		documentType,
		description: definition.description,
	};
}

function definitionsForDocumentType(
	documentType: EventHandlerDocumentType | undefined,
): readonly EventHandlerDefinition[] {
	switch (documentType) {
		case 'workbook':
			return WORKBOOK_EVENTS;
		case 'worksheet':
			return WORKSHEET_EVENTS;
		case 'chart':
			return CHART_EVENTS;
		default:
			return [];
	}
}

export const ALL_EVENT_DEFINITIONS: readonly EventHandlerDefinition[] = [
	...WORKBOOK_EVENTS,
	...WORKSHEET_EVENTS,
	...CHART_EVENTS,
];

function documentTypeForOwner(
	owner: EventHandlerDefinition['owner'],
): EventHandlerDocumentType | undefined {
	switch (owner) {
		case 'Workbook':
			return 'workbook';
		case 'Worksheet':
			return 'worksheet';
		case 'Chart':
			return 'chart';
	}
}

function inferDocumentType(moduleName: string | undefined): EventHandlerDocumentType {
	const lower = moduleName?.toLowerCase() ?? '';
	if (lower === 'thisworkbook') {
		return 'workbook';
	}
	if (/^chart\d*$/i.test(moduleName ?? '')) {
		return 'chart';
	}
	return 'worksheet';
}

function insideClosedProcedure(procedures: readonly ProcedureNode[], offset: number): boolean {
	return procedures.some(
		(proc) => proc.closed && offset >= proc.span.start && offset <= proc.span.end,
	);
}

function lineCompletionContext(source: string, offset: number): LineCompletionContext | undefined {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	const lineStart = currentLineStart(source, safeOffset);
	const prefix = source.slice(lineStart, safeOffset);
	const wordMatch = /[A-Za-z_][A-Za-z0-9_]*$/.exec(prefix);
	const currentWord = wordMatch?.[0] ?? '';
	const beforeWord = wordMatch ? prefix.slice(0, wordMatch.index) : prefix;

	if (/^\s*$/.test(beforeWord)) {
		return { currentWord, insertMode: 'fullProcedure' };
	}
	if (/^\s*Private\s+Sub\s+$/i.test(beforeWord)) {
		return { currentWord, insertMode: 'declarationTail' };
	}
	return undefined;
}

function currentLineStart(source: string, offset: number): number {
	const previousLf = source.lastIndexOf('\n', offset - 1);
	const previousCr = source.lastIndexOf('\r', offset - 1);
	return Math.max(previousLf, previousCr) + 1;
}

function toCompletion(
	definition: EventHandlerDefinition,
	insertMode: LineCompletionContext['insertMode'],
): EventHandlerCompletion {
	const signature = `${definition.name}(${definition.params})`;
	const insertText =
		insertMode === 'declarationTail'
			? `${signature}\n    $0\nEnd Sub`
			: `Private Sub ${signature}\n    $0\nEnd Sub`;
	return {
		name: definition.name,
		signature,
		detail: `${definition.owner} event handler`,
		documentation: definition.description,
		insertText,
	};
}
