import { describe, expect, it } from 'vitest';
import { incrementalModuleParse, incrementalModuleParseFromCache } from '../src/analyzer/parser/incrementalModuleParse';
import { parseModule, parseModuleFreshForTests } from '../src/analyzer/parser/parseModule';

const prefix = "Option Explicit\n" + "' Padding before the editable procedures\n".repeat(450);
const body = 'Public Sub First(ByVal argument As Long)\nDim value As Long\nvalue = 1\nDebug.Print value\nEnd Sub\n';
const tail = "''' Tail docs\nPublic Function Other() As String\nOther = \"hello\"\nEnd Function\n";

function compare(source: string, changed: string) {
    const previous = parseModuleFreshForTests(source);
    const original = JSON.stringify(previous);
    const fast = incrementalModuleParse(changed, source, previous, parseModuleFreshForTests);
    expect(fast).toBeDefined();
    expect(fast).toEqual(parseModuleFreshForTests(changed));
    expect(JSON.stringify(previous)).toBe(original);
    return fast!;
}

describe('incremental procedure parsing', () => {
    it.each(['value = 12', 'value = ', 'value = Left("x", 1)', 'value = α', 'Dim another As Double', 'If value Then', "' comment"])(
        'matches a full parse after body replacement with %s', replacement => {
            const source = prefix + body + tail;
            compare(source, source.replace('value = 1', replacement));
        });

    it.each(['\n', '\r\n', '\r'])('preserves all AST/diagnostic spans with %j lines', newline => {
        const source = (prefix + body + tail).replace(/\n/g, newline);
        compare(source, source.replace('Dim value As Long', 'Dim value As Double'));
    });

    it('reuses unchanged prefix nodes and rebases later procedure spans immutably', () => {
        const source = prefix + body + tail;
        const previous = parseModuleFreshForTests(source);
        const changed = source.replace('value = 1', 'value = 123');
        const fast = incrementalModuleParse(changed, source, previous, parseModuleFreshForTests)!;
        expect(fast.members[0]).toBe(previous.members[0]);
        expect(fast.members.at(-1)).not.toBe(previous.members.at(-1));
        expect(fast.members.at(-1)!.span.start).toBe(previous.members.at(-1)!.span.start + 2);
        expect(fast).toEqual(parseModuleFreshForTests(changed));
    });

    it.each([
        ['header', 'First(ByVal', 'Renamed(ByVal'],
        ['newline', 'value = 1', 'value = 1\nvalue = 2'],
        ['conditional directive', 'value = 1', '#If WIN64 Then'],
        ['terminator', 'End Sub', 'End Function'],
    ])('falls back on a %s edit', (_name, before, after) => {
        const source = prefix + body + tail;
        expect(incrementalModuleParse(source.replace(before, after), source, parseModuleFreshForTests(source), parseModuleFreshForTests)).toBeUndefined();
    });

    it('falls back for conditional or malformed procedures', () => {
        for (const text of [body.replace('value = 1', '#If WIN64 Then\nvalue = 1\n#End If'), body.replace('value = 1', 'If True Then\nvalue = 1')]) {
            const source = prefix + text + tail;
            expect(incrementalModuleParse(source.replace('value = 1', 'value = 12'), source, parseModuleFreshForTests(source), parseModuleFreshForTests)).toBeUndefined();
        }
    });

    it('probes only one compatible historical class before falling back for a module-level edit', () => {
        const source = prefix + body + tail;
        const first = parseModuleFreshForTests(source);
        let probed = 0;
        const snapshots = Array.from({ length: 8 }, (_, index) => ({
            source: source.replace('value = 1', 'value = ' + (index + 1)),
            module: { ...first, get members() { probed++; return first.members; } },
        }));
        const changed = prefix + 'unknown\n' + body + tail;
        expect(incrementalModuleParseFromCache(changed, snapshots, parseModuleFreshForTests)).toBeUndefined();
        expect(probed).toBe(0); // The newline gate rejected the only candidate before scanning members.
        const sameLineEdit = source.slice(0, prefix.length) + 'unknown ' + source.slice(prefix.length);
        expect(incrementalModuleParseFromCache(sameLineEdit, snapshots, parseModuleFreshForTests)).toBeUndefined();
        expect(probed).toBe(1);
        expect(parseModule(sameLineEdit)).toEqual(parseModuleFreshForTests(sameLineEdit));
    });

    it('matches fresh parses over a sequence of typing and Backspace changes', () => {
        let source = prefix + body + tail;
        parseModule(source);
        for (const value of ['12', '123', '12', '1', '', 'True', 'Left("text", 2)', 'Nothing', '1']) {
            const changed = source.replace(/value = [^\r\n]*/, 'value = ' + value);
            expect(parseModule(changed)).toEqual(parseModuleFreshForTests(changed));
            source = changed;
        }
    });
});
it.skipIf(!process.env.XLIDE_INCREMENTAL_PARSE_CORPUS)('compares actual large-class ASTs and parser work', async () => {
    const { readFileSync } = await import('node:fs');
    const original = readFileSync(process.env.XLIDE_INCREMENTAL_PARSE_CORPUS!, 'utf8');
    const source = original + '\nPublic Sub IncrementalProbe()\nDim n As Long\nn = 1\nEnd Sub\n';
    const previous = parseModuleFreshForTests(source);
    const full: number[] = [], fast: number[] = [];
    for (let i = 0; i < 9; i++) {
        const changed = source.replace('n = 1\nEnd Sub\n', 'n = ' + (100 + i) + '\nEnd Sub\n');
        // Warm full-source lexing so both samples isolate parser work.
        parseModuleFreshForTests(changed);
        let start = performance.now();
        const fresh = parseModuleFreshForTests(changed);
        full.push(performance.now() - start);
        start = performance.now();
        const incremental = incrementalModuleParse(changed, source, previous, parseModuleFreshForTests);
        fast.push(performance.now() - start);
        expect(incremental).toBeDefined();
        expect(incremental).toEqual(fresh);
    }
    process.stdout.write('Actual parser comparison: ' + JSON.stringify({ bytes: source.length, full, fast }) + '\n');
    const eligible = previous.members.filter(member => member.kind === 'Procedure' && member.closed &&
        member.body.length > 0 && !/^[ \t]*#/m.test(source.slice(member.span.start, member.span.end)) &&
        !previous.diagnostics.some(item => item.span.start < member.span.end && item.span.end >= member.span.start));
    const middle = eligible[Math.floor(eligible.length / 2)];
    expect(middle?.kind).toBe('Procedure');
    if (middle?.kind === 'Procedure') {
        const offset = middle.body[0].span.start;
        const changed = source.slice(0, offset) + ' ' + source.slice(offset);
        const start = performance.now();
        const incremental = incrementalModuleParse(changed, source, previous, parseModuleFreshForTests);
        process.stdout.write('Actual middle-procedure parser (ms): ' + (performance.now() - start) + '\n');
        expect(incremental).toBeDefined();
        expect(incremental).toEqual(parseModuleFreshForTests(changed));
    }
}, 30000);
