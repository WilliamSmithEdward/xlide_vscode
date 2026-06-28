// Pure helper for "hold errors on the current line" live-diagnostics behavior.
// Kept free of any runtime `vscode` dependency (type-only import) so it can be
// unit-tested without the VS Code module graph.
import type * as vscode from 'vscode';

/**
 * Returns the diagnostics that should be shown given the active cursor line:
 * "held while typing" diagnostics (syntax-category, e.g. an `If` with no `Then`
 * yet) whose range covers the active line are hidden until the cursor leaves the
 * line, matching the VBE which validates a line only once you leave it.
 *
 * When `activeLine` is undefined (the document is not the active editor) nothing
 * is suppressed.
 */
export function visibleDiagnosticsForActiveLine(
    diagnostics: readonly vscode.Diagnostic[],
    activeLine: number | undefined,
    isHeldWhileTyping: (diagnostic: vscode.Diagnostic) => boolean,
): vscode.Diagnostic[] {
    if (activeLine === undefined) {
        return [...diagnostics];
    }
    return diagnostics.filter((diagnostic) =>
        !(
            isHeldWhileTyping(diagnostic) &&
            diagnostic.range.start.line <= activeLine &&
            activeLine <= diagnostic.range.end.line
        ),
    );
}
