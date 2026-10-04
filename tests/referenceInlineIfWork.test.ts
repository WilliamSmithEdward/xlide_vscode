import { describe, expect, it } from 'vitest';
import { classifyReferenceKinds } from '../src/analyzer/references/referenceKinds';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

describe.each(['If c Then ', 'ElseIf c Then '])('nested reference tails (%s)', prefix => {
    it.each(['terminal', 'all'])('bounds token reads for %s requested names', mode => {
        const source = prefix.repeat(500) + "x = 1 ' " + mode;
        const tokens = tokenizeCached(source), target = source.lastIndexOf('x');
        let reads = 0;
        for (const token of tokens) {
            const raw = token.rawText;
            Object.defineProperty(token, 'rawText', { get() { reads++; return raw; } });
            Object.freeze(token);
        }
        Object.freeze(tokens);
        const offsets = mode === 'terminal' ? [target] : tokens.filter(token => token.kind === 'identifier').map(token => token.start);
        const result = classifyReferenceKinds(source, offsets);
        expect(result.size).toBe(offsets.length);
        expect([...result].every(([at, kind]) => kind === (at === target ? 'write' : 'read'))).toBe(true);
        expect(result.get(target)).toBe('write');
        expect(reads).toBeLessThan(20 * tokens.length);
    });
});

it.each([
    ['If a Then If b Then x = 1 Else y = 2 Else z = 3', ['x', 'y', 'z']],
    ['If a Then ElseIf b Then x = 1 Else y = 2', ['x', 'y']],
    ['Else If a Then x = 1 Else y = 2', ['x', 'y']],
    ['If a Then If (b) Then x = 1 Else y = 2', ['x', 'y']],
    ['If a Then If b x = 1 Else y = 2', ['y']],
    ['If a Then If b) Then x = 1 Else y = 2', []],
    ['If a Then If b( Then x = 1 Else y = 2', []],
    ['If a Then If b Then Dim x As Long Else Let y = 2', ['x', 'y']],
] as const)('preserves fragment writes for %s', (source, names) => {
    const offsets = tokenizeCached(source).filter(token => token.kind === 'identifier').map(token => token.start);
    const result = classifyReferenceKinds(source, offsets);
    expect([...result].filter(([, kind]) => kind === 'write').map(([at]) => source.slice(at).match(/^\w+/)![0])).toEqual(names);
});
