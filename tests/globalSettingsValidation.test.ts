import { describe, expect, it } from 'vitest';
import {
    normalizeXlideOptionExplicitSetting,
    validateXlideGlobalSettingsValues,
} from '../src/globalSettingsValidation';

const validSettings = {
    pythonPath: '',
    attachToRunningExcel: true,
    'diagnostics.enabled': true,
    'diagnostics.optionExplicit': 'warning',
    'analysis.visibleSeverities': ['error', 'warning', 'information'],
    'analysis.untrackedRules': [],
    'editor.blockLayout': 'comfy',
    'docs.enabled': true,
    'docs.metadataGlob': '**/*.vbref.xml',
};

describe('globalSettingsValidation', () => {
    it('accepts the extension default settings shape', () => {
        expect(validateXlideGlobalSettingsValues(validSettings)).toEqual([]);
    });

    it('reports malformed global analysis settings explicitly', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'analysis.visibleSeverities': ['error', 'hint'],
            'analysis.untrackedRules': ['option-explicit-missing', ''],
        })).toEqual([
            {
                key: 'xlide.analysis.visibleSeverities',
                message: 'Expected "xlide.analysis.visibleSeverities" entries to be one of: error, warning, information.',
                severity: 'error',
            },
            {
                key: 'xlide.analysis.untrackedRules',
                message: 'Expected "xlide.analysis.untrackedRules" entries to be non-empty strings.',
                severity: 'error',
            },
        ]);
    });

    it('reports malformed non-analysis settings through the same contract', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            attachToRunningExcel: 'yes',
            'diagnostics.optionExplicit': 'loud',
            'editor.blockLayout': 'spacious',
        }).map((problem) => problem.key)).toEqual([
            'xlide.attachToRunningExcel',
            'xlide.diagnostics.optionExplicit',
            'xlide.editor.blockLayout',
        ]);
    });

    it('does not expose hint as a diagnostic setting severity', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'diagnostics.optionExplicit': 'hint',
        }).map((problem) => problem.key)).toEqual(['xlide.diagnostics.optionExplicit']);
    });

    it('normalizes invalid Option Explicit severity to the default runtime value', () => {
        expect(normalizeXlideOptionExplicitSetting('off')).toBe('off');
        expect(normalizeXlideOptionExplicitSetting('error')).toBe('error');
        expect(normalizeXlideOptionExplicitSetting('loud')).toBe('warning');
        expect(normalizeXlideOptionExplicitSetting(undefined)).toBe('warning');
    });
});
