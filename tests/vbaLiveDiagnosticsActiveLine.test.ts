import { describe, it, expect } from 'vitest';
import type * as vscode from 'vscode';
import { visibleDiagnosticsForActiveLine } from '../src/vbaActiveLineDiagnostics';

// The active-line suppression only reads range.start.line / range.end.line and the
// held predicate, so plain shaped objects stand in for vscode.Diagnostic.
interface FakeDiag {
    id: string;
    range: { start: { line: number }; end: { line: number } };
}

function diag(id: string, startLine: number, endLine = startLine): FakeDiag {
    return { id, range: { start: { line: startLine }, end: { line: endLine } } };
}

const HELD_IDS = new Set(['if-missing-then', 'multiline-syntax']);
const isHeld = (d: vscode.Diagnostic): boolean => HELD_IDS.has((d as unknown as FakeDiag).id);

function visible(diags: FakeDiag[], activeLine: number | undefined): string[] {
    return visibleDiagnosticsForActiveLine(
        diags as unknown as vscode.Diagnostic[],
        activeLine,
        isHeld,
    ).map((d) => (d as unknown as FakeDiag).id);
}

describe('visibleDiagnosticsForActiveLine', () => {
    it('shows everything when no editor line is active', () => {
        expect(visible([diag('if-missing-then', 5), diag('unused-variable', 5)], undefined))
            .toEqual(['if-missing-then', 'unused-variable']);
    });

    it('hides a held (syntax) diagnostic on the cursor line, keeping others', () => {
        expect(visible([diag('if-missing-then', 5), diag('unused-variable', 5)], 5))
            .toEqual(['unused-variable']);
    });

    it('keeps a held diagnostic once the cursor is on a different line', () => {
        expect(visible([diag('if-missing-then', 5)], 6)).toEqual(['if-missing-then']);
    });

    it('hides a held multi-line diagnostic when the cursor is anywhere within it', () => {
        const diags = [diag('multiline-syntax', 4, 8)];
        expect(visible(diags, 4)).toEqual([]);
        expect(visible(diags, 6)).toEqual([]);
        expect(visible(diags, 8)).toEqual([]);
        expect(visible(diags, 9)).toEqual(['multiline-syntax']);
    });

    it('never hides a non-held (semantic) diagnostic, even on the cursor line', () => {
        expect(visible([diag('unused-variable', 5)], 5)).toEqual(['unused-variable']);
    });

    it('does not mutate the input array', () => {
        const diags = [diag('if-missing-then', 5)];
        visibleDiagnosticsForActiveLine(diags as unknown as vscode.Diagnostic[], 5, isHeld);
        expect(diags).toHaveLength(1);
    });
});
