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
        expect(got.every((item) => item.exclusive)).toBe(true);
    });

    it('supports typing xlide-test without the leading at sign', () => {
        const line = "' xlide-test-s";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual(['@xlide-test-skip']);
        expect(got[0].insertText).toBe('@xlide-test-skip reason="$1"');
        expect(got[0].range).toEqual({ start: 2, end: line.length });
    });

    it('turns a bare xlide prefix into a valid test directive comment', () => {
        const line = '    xli';
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual([
            '@xlide-test',
            '@xlide-test-skip',
            '@xlide-test-xfail',
        ]);
        expect(got[0].insertText).toBe("' @xlide-test");
        expect(got[0].range).toEqual({ start: 4, end: line.length });
        expect(got.every((item) => item.exclusive)).toBe(false);
    });

    it('turns a bare @ prefix into a valid test directive comment', () => {
        const line = '    @x';
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got[0].insertText).toBe("' @xlide-test");
        expect(got[0].range).toEqual({ start: 4, end: line.length });
    });

    it('expands the replacement range through the directive suffix after the cursor', () => {
        const line = "' @xlide-test";
        const got = resolveVbaTestDirectiveCompletions(line, "' @xlide".length);

        expect(got[0].range).toEqual({ start: 2, end: line.length });
    });

    it('offers metadata keys after a valid test directive', () => {
        const line = "' @xlide-test t";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual(['tags=', 'timeout=']);
        expect(got[0].insertText).toBe('tags=${1:smoke,fast}');
        expect(got[0].range).toEqual({ start: "' @xlide-test ".length, end: line.length });
        expect(got.every((item) => item.exclusive)).toBe(true);
    });

    it('offers reason metadata only for skip and expected-failure directives', () => {
        const skipLine = "' @xlide-test-skip r";
        const skip = resolveVbaTestDirectiveCompletions(skipLine, skipLine.length);
        const testLine = "' @xlide-test r";
        const test = resolveVbaTestDirectiveCompletions(testLine, testLine.length);

        expect(skip.map((item) => item.label)).toEqual(['reason=', 'requirement=']);
        expect(skip[0].insertText).toBe('reason="${1:Requires external workbook}"');
        expect(test.map((item) => item.label)).toEqual(['requirement=']);
    });

    it('offers exact and any-error expected-error metadata forms', () => {
        const line = "' @xlide-test expected";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual(['expected-error', 'expected-error=']);
        expect(got[0].insertText).toBe('expected-error');
        expect(got[1].insertText).toBe('expected-error=${1:13}');
    });

    it('offers deterministic value snippets for constrained metadata values', () => {
        const timeoutLine = "' @xlide-test timeout=";
        const timeout = resolveVbaTestDirectiveCompletions(timeoutLine, timeoutLine.length);
        const reasonLine = "' @xlide-test-xfail reason=\"";
        const reason = resolveVbaTestDirectiveCompletions(reasonLine, reasonLine.length);
        const expectedErrorLine = "' @xlide-test expected-error=";
        const expectedError = resolveVbaTestDirectiveCompletions(expectedErrorLine, expectedErrorLine.length);

        expect(timeout.map((item) => item.label)).toEqual(['10s', '30s', '2500ms']);
        expect(timeout[0].range).toEqual({ start: timeoutLine.length, end: timeoutLine.length });
        expect(reason.map((item) => item.label)).toEqual(['"Known issue pending fix"']);
        expect(reason[0].range).toEqual({ start: "' @xlide-test-xfail reason=".length, end: reasonLine.length });
        expect(expectedError.map((item) => item.label)).toEqual(['13', 'any']);
    });

    it('does not repeat completed metadata keys', () => {
        const line = "' @xlide-test tags=smoke t";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual(['timeout=']);
    });

    it('does not repeat bare expected-error metadata keys', () => {
        const line = "' @xlide-test expected-error e";
        const got = resolveVbaTestDirectiveCompletions(line, line.length);

        expect(got.map((item) => item.label)).toEqual([]);
    });

    it('does not offer directives in code, doc comments, quoted metadata, or non-test directives', () => {
        expect(resolveVbaTestDirectiveCompletions('Sub Test()', 'Sub Test()'.length)).toEqual([]);
        expect(resolveVbaTestDirectiveCompletions("''' @x", "''' @x".length)).toEqual([]);
        expect(resolveVbaTestDirectiveCompletions("' @xlide-test-skip reason=\"Known", "' @xlide-test-skip reason=\"Known".length)).toEqual([]);
        expect(resolveVbaTestDirectiveCompletions("' @xlide-analysis", "' @xlide-analysis".length)).toEqual([]);
    });
});
