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

	const documentType = ctx.documentType ?? inferDocumentType(ctx.moduleName);
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

function definitionsForDocumentType(
	documentType: EventHandlerDocumentType | undefined,
): readonly EventHandlerDefinition[] {
	switch (documentType) {
		case 'workbook':
			return WORKBOOK_EVENTS;
		case 'worksheet':
			return WORKSHEET_EVENTS;
		default:
			return [];
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
