/**
 * Shared CSS fragments for the XLIDE webviews. Emitted into each panel's
 * nonce-pinned <style> block; panels keep only genuinely panel-specific rules.
 */

export interface XlideAccentPaletteOptions {
    /** Surface color the accent mixes against; defaults to the editor background. */
    surface?: string;
    /** Full value for --xlide-accent-border. */
    accentBorder?: string;
}

/** The xlide accent palette variables; place inside a :root block. */
export function xlideAccentPaletteCss(options: XlideAccentPaletteOptions = {}): string {
    const surface = options.surface ?? 'var(--vscode-editor-background)';
    const accentBorder = options.accentBorder
        ?? 'color-mix(in srgb, var(--xlide-accent-blue) 78%, var(--vscode-panel-border))';
    return `--xlide-accent-blue: #2d5f94;
            --xlide-accent-blue-hover: #376fa8;
            --xlide-accent-background: color-mix(in srgb, var(--xlide-accent-blue) 82%, ${surface});
            --xlide-accent-hover-background: color-mix(in srgb, var(--xlide-accent-blue-hover) 84%, ${surface});
            --xlide-accent-border: ${accentBorder};`;
}

/** Base body reset; panels append a second body rule for their extras. */
export const WEBVIEW_BODY_CSS = `body {
            margin: 0;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
            font-size: var(--vscode-font-size);
        }`;

/** Accent-filled primary button with hover/disabled states. */
export const WEBVIEW_PRIMARY_BUTTON_CSS = `button {
            min-height: 32px;
            border: 1px solid var(--vscode-button-border, transparent);
            border-radius: 4px;
            padding: 5px 12px;
            color: var(--vscode-button-foreground);
            background: var(--xlide-accent-background);
            font: inherit;
            font-weight: 600;
            cursor: pointer;
        }
        button:hover:not(:disabled) {
            background: var(--xlide-accent-hover-background);
        }
        button:disabled {
            cursor: not-allowed;
            opacity: 0.55;
        }`;
