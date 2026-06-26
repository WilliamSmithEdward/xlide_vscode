// Single source of truth for the [xlide-vba] editor defaults. Declared in
// package.json configurationDefaults (kept in sync by vbaLanguageConfiguration
// tests) and enforced at activation by ensureXlideVbaEditorOverrides.
export const XLIDE_VBA_EDITOR_OVERRIDES: Array<{ key: string; value: boolean | number }> = [
    { key: 'detectIndentation', value: false },
    { key: 'tabSize', value: 4 },
    { key: 'minimap.enabled', value: true },
    { key: 'minimap.renderCharacters', value: false },
    { key: 'minimap.showMarkSectionHeaders', value: false },
    { key: 'minimap.showRegionSectionHeaders', value: false },
    { key: 'overviewRulerBorder', value: false },
    { key: 'overviewRulerLanes', value: 3 },
    // Make XLIDE's workbook-aware suggestion list the primary completion UI for
    // VBA modules: turn off ghost-text inline suggestions (and the inline suggest
    // preview) so they don't pre-empt the dropdown. Users who want inline
    // suggestions in VBA can re-enable them with a `[xlide-vba]` language override.
    { key: 'inlineSuggest.enabled', value: false },
    { key: 'suggest.preview', value: false },
];
