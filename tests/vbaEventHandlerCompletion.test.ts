import { describe, expect, it } from 'vitest';
import {
	EventHandlerCompletionContext,
	resolveEventHandlerCompletions,
} from '../src/analyzer';

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
		expect(item?.insertText).toBe('Private Sub Workbook_Open()\n    $0\nEnd Sub');
	});

	it('inserts chart event handler stubs from chart document modules', () => {
		const src = 'Option Explicit\nChart';
		const item = completions(src, 'Chart', {
			moduleName: 'Chart1',
			moduleKind: 'document',
			documentType: 'chart',
		}).find((completion) => completion.name === 'Chart_MouseDown');
		expect(item?.insertText).toBe(
			'Private Sub Chart_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)\n    $0\nEnd Sub',
		);
	});

	it('inserts only the declaration tail after an existing Private Sub prefix', () => {
		const src = 'Private Sub Work';
		const item = completions(src, 'Work', {
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			documentType: 'workbook',
		}).find((completion) => completion.name === 'Workbook_Open');
		expect(item?.insertText).toBe('Workbook_Open()\n    $0\nEnd Sub');
	});
});
