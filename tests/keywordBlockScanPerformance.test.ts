import { afterEach, describe, expect, it, vi } from 'vitest';
import * as smartEnter from '../src/vbaSmartEnter';
import { resolveKeywordCompletions } from '../src/analyzer/completion/keywordCompletion';

afterEach(() => vi.restoreAllMocks());

describe('keyword block scan work', () => {
    it('does not scan enclosing blocks for an ordinary identifier prefix', () => {
        const source = 'Sub Padding()\nDebug.Print 1\nEnd Sub\n'.repeat(1200) +
            'Sub Probe()\nIf ready Then\nLatencyValue';
        const scan = vi.spyOn(smartEnter, 'openSmartBlockClosersBefore');
        expect(resolveKeywordCompletions(source, source.length)).toEqual({ items: [], exclusive: false });
        expect(scan).not.toHaveBeenCalled();
    });

    it('scans once for empty prefixes while retaining the active closer and branch rows', () => {
        const source = 'Sub BlockScanWorkProbe()\nIf ready Then\n    ';
        const scan = vi.spyOn(smartEnter, 'openSmartBlockClosersBefore');
        const labels = resolveKeywordCompletions(source, source.length).items.map(item => item.label);
        expect(labels).toContain('End If');
        expect(labels).toContain('ElseIf');
        expect(labels).toContain('Else');
        expect(scan).toHaveBeenCalledTimes(1);
    });

    it('reuses enclosing blocks across partial-word edits and refreshes preceding block changes', () => {
        const prefix = 'Sub ChangingKeywordProbe()\nIf ready Then\n    ';
        const scan = vi.spyOn(smartEnter, 'openSmartBlockClosersBefore');
        for (const partial of ['', 'e', 'en']) {
            expect(resolveKeywordCompletions(prefix + partial, prefix.length + partial.length).items
                .map(item => item.label)).toContain('End If');
        }
        expect(scan).toHaveBeenCalledTimes(1);
        const changed = prefix.replace('If ready Then', 'While ready');
        expect(resolveKeywordCompletions(changed, changed.length).items.map(item => item.label)).toContain('Wend');
        expect(scan).toHaveBeenCalledTimes(2);
    });

    it.each([
        ['If ready Then', 'el', 'Else'],
        ['Select Case value', 'ca', 'Case'],
        ['Do', 'lo', 'Loop Until'],
        ['For item = 1 To 10', 'nextitem', 'Next item'],
        ['While ready', 'we', 'Wend'],
    ])('retains context rows for %s with prefix %s', (opener, prefix, expected) => {
        const source = 'Sub Probe()\n' + opener + '\n    ' + prefix;
        expect(resolveKeywordCompletions(source, source.length).items.map(item => item.label)).toContain(expected);
    });
});
