import { describe, expect, it } from 'vitest';
import { resolveVbaTestDirectiveCompletions } from '../src/vbaTestDirectiveCompletion';

describe('VBA test directive completion', () => {
    it('offers @xlide-test directives from apostrophe comments', () => {
        const line = "' @x";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual([
            '@xlide-test',
            '@xlide-test-skip',
            '@xlide-test-xfail',
        ]);
        expect(got[0].range).toEqual({ start: 2, end: 4 });
    });

    it('supports typing xlide-test without the leading at sign', () => {
        const line = "' xlide-test-s";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual(['@xlide-test-skip']);
        expect(got[0].insertText).toBe('@xlide-test-skip reason="$1"');
        expect(got[0].range).toEqual({ start: 2, end: line.length });
    });

    it('expands the replacement range through the directive suffix after the cursor', () => {
        const line = "' @xlide-test";
        const got = resolveVbaTestDirectiveCompletions(line, "' @xlide".length);

        expect(got[0].range).toEqual({ start: 2, end: line.length });
    });

    it('does not offer directives in code, doc comments, or directive metadata', () => {
        expect(resolveVbaTestDirectiveCompletions('Sub Test()', 'Sub Test()'.length)).toEqual([]);
        expect(resolveVbaTestDirectiveCompletions("''' @x", "''' @x".length)).toEqual([]);
        expect(resolveVbaTestDirectiveCompletions("' @xlide-test tags=", "' @xlide-test tags=".length)).toEqual([]);
        expect(resolveVbaTestDirectiveCompletions("' @xlide-analysis", "' @xlide-analysis".length)).toEqual([]);
    });
});
