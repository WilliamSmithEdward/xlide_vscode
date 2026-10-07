import { afterEach, expect, it, vi } from 'vitest';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { macroNameStringAt } from '../src/analyzer/completion/macroNames';

afterEach(() => vi.restoreAllMocks());

it.each(['code', 'macro'])('bounds token lookups for %s near the end of a large module', shape => {
    const tail = shape === 'macro' ? 'obj.OnAction = "Demo.Run"' : 'ThisWorkbook.Sheets(1).ce';
    const source = 'Sub Demo()\n' + 'value = value + 1\n'.repeat(10000) + tail + '\nEnd Sub';
    const tokens = lexer.tokenizeCached(source);
    let reads = 0;
    const counted = new Proxy(tokens, {
        get(target, key, receiver) {
            if (typeof key === 'string' && /^\d+$/.test(key)) { reads++; }
            return Reflect.get(target, key, receiver);
        },
    });
    vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(counted);
    const offset = source.indexOf(tail) + tail.length - (shape === 'macro' ? 1 : 0);
    const result = macroNameStringAt(source, offset);
    if (shape === 'macro') {
        expect(result?.text).toBe('Demo.Run');
    } else {
        expect(result).toBeUndefined();
    }
    expect(reads).toBeLessThan(100);
});

it('keeps the string cursor boundaries when adjacent tokens share an offset', () => {
    const source = 'obj.OnAction="Demo.Run"';
    const open = source.indexOf('"');
    expect(macroNameStringAt(source, open)).toBeUndefined();
    expect(macroNameStringAt(source, open + 1)?.text).toBe('Demo.Run');
    expect(macroNameStringAt(source, source.length)?.text).toBe('Demo.Run');
    expect(macroNameStringAt(source, source.length + 1)).toBeUndefined();
});
