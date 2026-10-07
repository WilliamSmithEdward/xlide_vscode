import { describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => {
	const { vscodeMock, Range } = await import('./helpers/vscodeMock');
	return vscodeMock({ Range,
		Color: class { constructor(public red: number, public green: number, public blue: number, public alpha: number) {} },
		ColorInformation: class { constructor(public range: unknown, public color: unknown) {} },
		ColorPresentation: class { constructor(public label: string) {} },
		TextEdit: { replace: (range: unknown, newText: string) => ({ range, newText }) },
	});
});

import * as vscode from 'vscode';
import { VbaColorProvider } from '../src/vbaColorProvider';

function document(source: string) {
	return {
		getText: () => source,
		positionAt: (offset: number) => {
			const lines = source.slice(0, offset).split('\n');
			return new vscode.Position(lines.length - 1, lines.at(-1)!.length);
		},
		offsetAt: (position: vscodeTypes.Position) => source.split('\n').slice(0, position.line)
			.reduce((length, line) => length + line.length + 1, 0) + position.character,
	} as unknown as vscodeTypes.TextDocument;
}

describe('native color provider', () => {
	it('returns the exact editable expression range and opaque color', () => {
		const doc = document('Me.BackColor = RGB(255, 0, 0)');
		const [info] = new VbaColorProvider().provideDocumentColors(doc, { isCancellationRequested: false } as vscodeTypes.CancellationToken);
		expect(info.range.start.character).toBe(15);
		expect(info.range.end.character).toBe(29);
		expect(info.color).toMatchObject({ red: 1, green: 0, blue: 0, alpha: 1 });
	});
	it('writes safe picker edits when RGB is shadowed', () => {
		const doc = document('Sub T()\nDim RGB As Long\nMe.BackColor = vbRed\nEnd Sub');
		const range = new vscode.Range(new vscode.Position(2, 15), new vscode.Position(2, 20));
		const results = new VbaColorProvider().provideColorPresentations(new vscode.Color(0, 0, 1, 1), { document: doc, range });
		expect(results[0].label).toBe('VBA.RGB(0, 0, 255)');
		expect(results[0].textEdit?.range).toBe(range);
		expect(results[0].textEdit?.newText).toBe(results[0].label);
	});
	it('uses numeric presentations when both RGB and VBA are shadowed', () => {
		const doc = document('Sub T()\nDim RGB As Long, VBA As Long\nMe.BackColor = vbRed\nEnd Sub');
		const range = new vscode.Range(new vscode.Position(2, 15), new vscode.Position(2, 20));
		expect(new VbaColorProvider().provideColorPresentations(new vscode.Color(1, 0, 0, 1), { document: doc, range })
			.map(p => p.label)).toEqual(['&H0000FF&', '255']);
	});
	it('does not read a document for an already-cancelled request', () => {
		const doc = { getText: vi.fn() } as unknown as vscodeTypes.TextDocument;
		expect(new VbaColorProvider().provideDocumentColors(doc, { isCancellationRequested: true } as vscodeTypes.CancellationToken)).toEqual([]);
		expect(doc.getText).not.toHaveBeenCalled();
	});
});
