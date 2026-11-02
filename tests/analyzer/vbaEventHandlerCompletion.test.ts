import { describe, expect, it } from 'vitest';
import {
	EventHandlerCompletionContext,
	resolveEventHandlerCompletions,
} from '../../src/analyzer';
import { getAccessObjectModel } from '../../src/analyzer/host/accessObjectModel';
import { OFFICE_HOST_APPS, type OfficeHostApp } from '../../src/officeHostApps';

function at(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function completions(
	src: string,
	marker: string,
	ctx: EventHandlerCompletionContext,
) {
	return resolveEventHandlerCompletions(src, at(src, marker), ctx);
}

function names(
	src: string,
	marker: string,
	ctx: EventHandlerCompletionContext,
): string[] {
	return completions(src, marker, ctx).map((item) => item.name);
}

describe('event-handler completion', () => {
	it('offers workbook event handlers in ThisWorkbook document modules', () => {
		const src = 'Option Explicit\nWork\n';
		const got = names(src, 'Work', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		});
		expect(got).toContain('Workbook_Open');
		expect(got).toContain('Workbook_BeforeClose');
		expect(got).not.toContain('Worksheet_Change');
	});

	it('offers worksheet event handlers in worksheet document modules', () => {
		const src = 'Option Explicit\nWorksheet_\n';
		const got = names(src, 'Worksheet_', {
			moduleName: 'Sheet1',
			moduleKind: 'document',
			documentType: 'worksheet',
		});
		expect(got).toContain('Worksheet_Change');
		expect(got).toContain('Worksheet_SelectionChange');
		expect(got).not.toContain('Workbook_Open');
	});

	it('does not offer workbook or worksheet event handlers outside document modules', () => {
		const src = 'Work\n';
		const got = names(src, 'Work', {
			moduleName: 'Module1',
			moduleKind: 'standard',
		});
		expect(got).toEqual([]);
	});

	it('offers chart event handlers in chart document modules', () => {
		const src = 'Chart_\n';
		const got = names(src, 'Chart_', {
			moduleName: 'Chart1',
			moduleKind: 'document',
			documentType: 'chart',
		});
		expect(got).toContain('Chart_Calculate');
		expect(got).toContain('Chart_MouseDown');
		expect(got).toContain('Chart_SeriesChange');
		expect(got).toContain('Chart_BeforeDoubleClick');
		expect(got).not.toContain('Worksheet_Change');
		expect(got).not.toContain('Workbook_Open');
	});

	it('does not offer worksheet handlers for chart document modules', () => {
		const src = 'Worksheet_\n';
		const got = names(src, 'Worksheet_', {
			moduleName: 'Chart1',
			moduleKind: 'document',
			documentType: 'chart',
		});
		expect(got).toEqual([]);
	});

	it('does not suggest a handler already declared in the module', () => {
		const src =
			'Private Sub Workbook_Open()\n' +
			'End Sub\n' +
			'\n' +
			'Work\n';
		const got = names(src, '\nWork', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		});
		expect(got).not.toContain('Workbook_Open');
		expect(got).toContain('Workbook_BeforeClose');
	});

	it('does not suggest an existing chart handler in a chart module', () => {
		const src =
			'Private Sub Chart_Calculate()\n' +
			'End Sub\n' +
			'\n' +
			'Chart_\n';
		const got = names(src, '\nChart_', {
			moduleName: 'Chart1',
			moduleKind: 'document',
			documentType: 'chart',
		});
		expect(got).not.toContain('Chart_Calculate');
		expect(got).toContain('Chart_Activate');
	});

	it('does not offer event stubs inside an existing procedure body', () => {
		const src = 'Sub Test()\n    Work\nEnd Sub\n';
		const got = names(src, 'Work', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		});
		expect(got).toEqual([]);
	});

	it('inserts a full Private Sub stub from a blank module-level line', () => {
		const src = 'Option Explicit\nWork';
		const item = completions(src, 'Work', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		}).find((completion) => completion.name === 'Workbook_Open');
		expect(item?.insertText).toBe('Private Sub Workbook_Open()\n\t$0\nEnd Sub');
	});

	it('inserts chart event handler stubs from chart document modules', () => {
		const src = 'Option Explicit\nChart';
		const item = completions(src, 'Chart', {
			moduleName: 'Chart1',
			moduleKind: 'document',
			documentType: 'chart',
		}).find((completion) => completion.name === 'Chart_MouseDown');
		expect(item?.insertText).toBe(
			'Private Sub Chart_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)\n\t$0\nEnd Sub',
		);
	});

	it('inserts only the declaration tail after an existing Private Sub prefix', () => {
		const src = 'Private Sub Work';
		const item = completions(src, 'Work', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		}).find((completion) => completion.name === 'Workbook_Open');
		expect(item?.insertText).toBe('Workbook_Open()\n\t$0\nEnd Sub');
	});
});

describe('Word document event handlers (issue #25)', () => {
	it('offers Document_* handlers in a ThisDocument module', () => {
		const src = 'Option Explicit\nDoc\n';
		const got = names(src, 'Doc', {
			moduleName: 'ThisDocument',
			moduleKind: 'document',
			documentType: 'document',
		});
		expect(got).toContain('Document_Open');
		expect(got).toContain('Document_New');
		expect(got).toContain('Document_Close');
		expect(got).toContain('Document_ContentControlOnEnter');
		expect(got).not.toContain('Workbook_Open');
		expect(got).not.toContain('Worksheet_Change');
	});

	it('infers the document type from the ThisDocument name when unstated', () => {
		const src = 'Option Explicit\nDoc\n';
		const got = names(src, 'Doc', {
			moduleName: 'ThisDocument',
			moduleKind: 'document',
		});
		expect(got).toContain('Document_Open');
		expect(got).not.toContain('Worksheet_Change');
	});

	it('never offers Document_* handlers in Excel document modules', () => {
		const src = 'Option Explicit\nDoc\n';
		const got = names(src, 'Doc', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		});
		expect(got).not.toContain('Document_Open');
	});
});

describe('where a handler stub is offered', () => {
	// One handler per Office host, typed far enough to name it.
	const hosts: readonly (readonly [string, EventHandlerCompletionContext, string, string])[] = [
		['Excel ThisWorkbook', { moduleName: 'ThisWorkbook', moduleKind: 'document', documentType: 'workbook', host: 'excel' }, 'Workbook_O', 'Workbook_Open'],
		['Word ThisDocument', { moduleName: 'ThisDocument', moduleKind: 'document', documentType: 'document', host: 'word' }, 'Document_O', 'Document_Open'],
		['a PowerPoint UserForm', { moduleName: 'UserForm1', moduleKind: 'userform', host: 'powerpoint' }, 'UserForm_I', 'UserForm_Initialize'],
		['an Access form', { moduleName: 'Form_Orders', moduleKind: 'userform', host: 'access', meType: 'Access.Form', model: getAccessObjectModel() }, 'Form_Loa', 'Form_Load'],
		['a class module', { moduleName: 'Account', moduleKind: 'class', host: 'excel' }, 'Class_T', 'Class_Terminate'],
	];
	const atEnd = (src: string, ctx: EventHandlerCompletionContext) =>
		resolveEventHandlerCompletions(src, src.length, ctx).map((item) => item.name);

	it.each(hosts)('offers none in an unfinished procedure at the end of the module, or in a Type or Enum (%s)', (_, ctx, typed, handler) => {
		for (const eol of ['\n', '\r\n']) {
			expect(atEnd(`Option Explicit${eol}Sub Work()${eol}${typed}`, ctx)).toEqual([]);
			expect(atEnd(`Option Explicit${eol}Private Enum Kind${eol}${typed}`, ctx)).toEqual([]);
			expect(atEnd(`Option Explicit${eol}Private Type Point${eol}${typed}`, ctx)).toEqual([]);
			expect(names(`Private Enum Kind${eol}${typed}${eol}End Enum${eol}`, `${eol}${typed}`, ctx)).toEqual([]);
			expect(names(`Private Type Point${eol}${typed}${eol}End Type${eol}`, `${eol}${typed}`, ctx)).toEqual([]);
			expect(atEnd(`Option Explicit${eol}Sub Work()${eol}End Sub${eol}${typed}`, ctx)).toEqual([handler]);
		}
	});

	it.each(hosts)('offers none on a row that continues the line above (%s)', (_, ctx, typed, handler) => {
		for (const eol of ['\n', '\r\n']) {
			expect(atEnd(`Option Explicit${eol}Const Limit = _${eol}${typed}`, ctx)).toEqual([]);
			expect(atEnd(`Option Explicit${eol}' a note _${eol}${typed}`, ctx)).toEqual([]);
			expect(atEnd(`Option Explicit${eol}Const Limit = 1${eol}${typed}`, ctx)).toEqual([handler]);
		}
	});

	it.each(hosts)('completes the name on an unfinished header row, whatever follows it (%s)', (_, ctx, typed, handler) => {
		for (const after of ['', '\n', '\r\n', '\n\nPrivate Sub Work()\nEnd Sub\n']) {
			const tail = completions(`Option Explicit\nPrivate Sub ${typed}${after}`, `Sub ${typed}`, ctx);
			expect(tail.map((item) => item.name)).toEqual([handler]);
			expect(tail[0].insertText.startsWith(`${handler}(`)).toBe(true);
		}
		// The row below the unfinished header is still its body.
		expect(names(`Private Sub Work()\n${typed}\n`, `\n${typed}`, ctx)).toEqual([]);
	});
});

describe('class module event handlers', () => {
	const hosts = Object.keys(OFFICE_HOST_APPS) as OfficeHostApp[];

	it.each(hosts)('offers Class_Initialize and Class_Terminate in a %s class module', (host) => {
		const src = 'Option Explicit\nPrivate m As Long\nC';
		expect(completions(src, '\nC', { moduleName: 'Account', moduleKind: 'class', host })).toEqual([
			{
				name: 'Class_Initialize',
				signature: 'Class_Initialize()',
				detail: 'Class event handler',
				documentation: 'Occurs when an instance of the class is created.',
				insertText: 'Private Sub Class_Initialize()\n\t$0\nEnd Sub',
			},
			{
				name: 'Class_Terminate',
				signature: 'Class_Terminate()',
				detail: 'Class event handler',
				documentation: 'Occurs when all references to an instance of the class are removed from memory.',
				insertText: 'Private Sub Class_Terminate()\n\t$0\nEnd Sub',
			},
		]);
	});

	it('offers them in a VB6 class module too', () => {
		expect(names('Option Explicit\nClass_', 'Class_', { moduleName: 'Account', moduleKind: 'class', host: 'vb6' }))
			.toEqual(['Class_Initialize', 'Class_Terminate']);
	});

	it('offers both on an empty module-level line and only the tail after Private Sub', () => {
		const ctx: EventHandlerCompletionContext = { moduleName: 'Account', moduleKind: 'class' };
		expect(names('Option Explicit\n', 'Explicit\n', ctx)).toEqual(['Class_Initialize', 'Class_Terminate']);
		const tail = completions('Private Sub Class_T', 'Class_T', ctx);
		expect(tail.map((item) => item.insertText)).toEqual(['Class_Terminate()\n\t$0\nEnd Sub']);
	});

	it('leaves out a handler the class already has, and offers none inside a procedure', () => {
		const ctx: EventHandlerCompletionContext = { moduleName: 'Account', moduleKind: 'class' };
		const src = 'Private Sub class_initialize()\nEnd Sub\nClass_\nSub Work()\n    Class_\nEnd Sub\n';
		expect(names(src, 'End Sub\nClass_', ctx)).toEqual(['Class_Terminate']);
		expect(names(src, '    Class_', ctx)).toEqual([]);
	});

	it.each([
		['an Excel workbook module', { moduleName: 'ThisWorkbook', moduleKind: 'document', documentType: 'workbook', host: 'excel' }],
		['an Excel worksheet module', { moduleName: 'Sheet1', moduleKind: 'document', documentType: 'worksheet', host: 'excel' }],
		['a Word document module', { moduleName: 'ThisDocument', moduleKind: 'document', documentType: 'document', host: 'word' }],
		['a PowerPoint UserForm', { moduleName: 'UserForm1', moduleKind: 'userform', host: 'powerpoint' }],
		['an Access form module', { moduleName: 'Form_Orders', moduleKind: 'userform', host: 'access', meType: 'Access.Form' }],
		['an Access report module', { moduleName: 'Report_Orders', moduleKind: 'userform', host: 'access', meType: 'Access.Report' }],
		['a standard module', { moduleName: 'Module1', moduleKind: 'standard', host: 'excel' }],
	] as const)('offers no Class_ handler in %s', (_, ctx) => {
		expect(names('Option Explicit\nClass_', 'Class_', ctx)).toEqual([]);
	});
});
