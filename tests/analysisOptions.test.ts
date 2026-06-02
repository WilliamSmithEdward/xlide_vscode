import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
    workspace: {
        getConfiguration: () => ({
            get: () => [],
            update: async () => undefined,
        }),
    },
    ConfigurationTarget: {
        Global: true,
    },
}));

import { setAnalysisRuleTrackedInList } from '../src/analysisOptions';

describe('analysis rule tracking options', () => {
    it('adds a rule to the untracked list when tracking is disabled', () => {
        expect(setAnalysisRuleTrackedInList([], 'Option-Explicit-Missing', false))
            .toEqual(['option-explicit-missing']);
    });

    it('removes a rule from the untracked list when tracking is enabled', () => {
        expect(setAnalysisRuleTrackedInList(
            ['argument-count', 'option-explicit-missing'],
            'OPTION-EXPLICIT-MISSING',
            true,
        )).toEqual(['argument-count']);
    });

    it('normalizes and deduplicates the untracked list', () => {
        expect(setAnalysisRuleTrackedInList(
            [' option-explicit-missing ', 'OPTION-EXPLICIT-MISSING'],
            'option-explicit-missing',
            false,
        )).toEqual(['option-explicit-missing']);
    });
});
