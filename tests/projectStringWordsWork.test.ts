import { afterEach, describe, expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import * as tokenHelpers from '../src/analyzer/lexer/tokenHelpers';

function moduleWithLiteral(moduleName: string, literal: string) {
    return { moduleName, moduleKind: 'standard' as const,
        source: `Sub Demo()\nDim value As String\nvalue = "${literal}"\nEnd Sub\n` };
}
afterEach(() => vi.restoreAllMocks());

describe('lazy project string-word facts', () => {
    it('does no string-word extraction until that query is needed', () => {
        const words = vi.spyOn(tokenHelpers, 'identifierWords');
        const index = new ProjectIndex();
        for (let i = 0; i < 80; i++) { index.setModule(moduleWithLiteral(`Module${i}`, `word${i}`)); }
        expect(index.moduleNames()).toHaveLength(80);
        expect(words).not.toHaveBeenCalled();
        expect(index.stringLiteralWords().size).toBe(80);
        expect(words).toHaveBeenCalledTimes(80);
        expect(index.stringLiteralWords().size).toBe(80);
        expect(words).toHaveBeenCalledTimes(80);
    });

    it('keeps unchanged facts and invalidates only replacement/deleted modules', () => {
        const words = vi.spyOn(tokenHelpers, 'identifierWords');
        const index = new ProjectIndex();
        index.setModule(moduleWithLiteral('First', 'alpha beta'));
        index.setModule(moduleWithLiteral('Second', 'gamma'));
        const first = index.stringLiteralWords();
        expect([...first]).toEqual(['alpha', 'beta', 'gamma']);
        expect(words).toHaveBeenCalledTimes(2);
        words.mockClear();
        index.setModule(moduleWithLiteral('FIRST', 'delta'));
        expect(words).not.toHaveBeenCalled();
        expect([...index.stringLiteralWords()]).toEqual(['delta', 'gamma']);
        expect(words.mock.calls.map(call => call[0])).toEqual(['"delta"']);
        expect([...first]).toEqual(['alpha', 'beta', 'gamma']);
        words.mockClear();
        index.removeModule('Second');
        expect([...index.stringLiteralWords()]).toEqual(['delta']);
        expect(words).not.toHaveBeenCalled();
        index.setModule(moduleWithLiteral('Second', 'epsilon'));
        expect([...index.stringLiteralWords()]).toEqual(['delta', 'epsilon']);
        expect(words.mock.calls.map(call => call[0])).toEqual(['"epsilon"']);
    });

    it('keeps literal facts when an existing source is submitted with new metadata', () => {
        const words = vi.spyOn(tokenHelpers, 'identifierWords');
        const index = new ProjectIndex();
        const input = moduleWithLiteral('Stable', 'alpha');
        index.setModule(input);
        expect([...index.stringLiteralWords()]).toEqual(['alpha']);
        words.mockClear();
        index.setModule({ ...input, moduleName: 'STABLE', moduleKind: 'class', predeclaredId: true,
            source: (' ' + input.source).slice(1) });
        expect(index.moduleNames()).toEqual(['STABLE']);
        expect([...index.stringLiteralWords()]).toEqual(['alpha']);
        expect(words).not.toHaveBeenCalled();
    });

    it('retains empty per-module results across unrelated index changes', () => {
        const words = vi.spyOn(tokenHelpers, 'identifierWords');
        const index = new ProjectIndex();
        index.setModule({ moduleName: 'Empty', moduleKind: 'standard', source: 'Sub Demo()\nEnd Sub\n' });
        expect([...index.stringLiteralWords()]).toEqual([]);
        index.setModule(moduleWithLiteral('Other', 'hello'));
        expect([...index.stringLiteralWords()]).toEqual(['hello']);
        expect(words).toHaveBeenCalledTimes(1);
        index.removeModule('Other');
        expect([...index.stringLiteralWords()]).toEqual([]);
        expect(words).toHaveBeenCalledTimes(1);
    });

    it('preserves Unicode, escaped strings, attributes and comment exclusion', () => {
        const index = new ProjectIndex();
        index.setModule({ moduleName: 'Text', moduleKind: 'standard', source:
            'Attribute VB_Name = "Text"\nSub Demo()\nvalue = "Refresh ""poll"" ไทย 名字"\n\' "ignored"\nRem "hidden"\nEnd Sub\n' });
        expect([...index.stringLiteralWords()]).toEqual(['text', 'refresh', 'poll', 'ไทย', '名字']);
    });
});
