import { describe, expect, it } from 'vitest';
import { setAnalysisRuleTrackedInList } from '../src/analysisSettingsCore';

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
