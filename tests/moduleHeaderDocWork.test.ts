import { afterEach, describe, expect, it, vi } from 'vitest';
import { extractModuleHeaderDoc } from '../src/analyzer/docs/docComment';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { parseModule } from '../src/analyzer/parser/parseModule';
import * as sourceScan from '../src/vbaSourceScan';

afterEach(() => vi.restoreAllMocks());
const summary = "''' <summary>Module summary.</summary>";
const remarks = "''' <remarks>Header details.</remarks>";

describe('bounded module header documentation', () => {
    it.each(['\n', '\r\n', '\r'])('preserves start-offset behavior with %j lines', eol => {
        const prefix = 'Attribute VB_Name = "M"' + eol;
        const source = prefix + [summary, remarks, '', 'Option Explicit', 'Sub Demo()', 'End Sub'].join(eol);
        expect(extractModuleHeaderDoc(source)).toBeUndefined();
        for (const offset of [prefix.length - 1, prefix.length, prefix.length - 0.5]) {
            expect(extractModuleHeaderDoc(source, offset)?.summary).toBe('Module summary.');
            expect(extractModuleHeaderDoc(source, offset)?.remarks).toBe('Header details.');
        }
        // An offset inside a physical line starts the scan at the next line.
        const fromRemarks = extractModuleHeaderDoc(source, prefix.length + 1);
        expect(fromRemarks?.summary).toBeUndefined();
        expect(fromRemarks?.remarks).toBe('Header details.');
        for (const offset of [NaN, Infinity, source.length, source.length + 1]) {
            expect(extractModuleHeaderDoc(source, offset)).toBeUndefined();
        }
        const bare = summary + eol;
        for (const offset of [-Infinity, -1, 0]) expect(extractModuleHeaderDoc(bare, offset)?.summary).toBe('Module summary.');
        expect(extractModuleHeaderDoc(summary)?.summary).toBe('Module summary.');
    });

    it.each(['', "' ordinary comment", 'Option Explicit', 'option compare text'])('retains module ownership before boundary %j', boundary => {
        const source = [summary, remarks, boundary, 'Sub Demo()', 'End Sub'].join('\n');
        expect(extractModuleHeaderDoc(source)?.summary).toBe('Module summary.');
    });

    it.each(['Sub Demo()', "' @xlide-test", "'' @xlide-analysis-disable-next-member unused-procedure"])('leaves adjacent member/directive documentation unclaimed before %j', boundary => {
        const source = [summary, boundary, 'Sub Demo()', 'End Sub'].join('\n');
        expect(extractModuleHeaderDoc(source)).toBeUndefined();
    });

    it('stops before a large class body rather than reading every physical line', () => {
        const prefix = 'Attribute VB_Name = "M"\r\n';
        const source = prefix + [summary, remarks, '', 'Sub Demo()', "' large body\r\n".repeat(26000), 'End Sub'].join('\r\n');
        const read = vi.spyOn(sourceScan, 'lineEndAtOrAfter');
        expect(extractModuleHeaderDoc(source, prefix.length)?.summary).toBe('Module summary.');
        expect(read).toHaveBeenCalledTimes(4);
    });

    it('keeps symbol building bounded to the header after a large body edit', () => {
        const source = ['Attribute VB_Name = "M"', summary, '', 'Sub Demo()', "' large body\n".repeat(26000), 'End Sub'].join('\n');
        for (const text of [source, source + ' ']) {
            parseModule(text);
            const read = vi.spyOn(sourceScan, 'lineEndAtOrAfter');
            expect(buildModuleSymbols('M', 'standard', text).root.doc?.summary).toBe('Module summary.');
            expect(read.mock.calls.length).toBeLessThanOrEqual(3);
            read.mockRestore();
        }
    });
});
