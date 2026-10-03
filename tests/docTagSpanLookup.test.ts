import { describe, expect, it } from 'vitest';
import { leadingDocLines, scanDocTags } from '../src/analyzer/docs/docComment';

describe('documentation tag source mapping', () => {
    it.each(['\n', '\r\n'])('keeps physical spans across directives and multiline tags (%j)', (eol) => {
        const source = [
            "''' <summary>Intro",
            "' @xlide-analysis-disable DOC001",
            "''' <param",
            "''' name=\"A&amp;B\" type=\"Long\">Value</param>",
            "''' </summary><remarks />",
            "''' <param name=\"Earlier\" name=\"Last\" />",
            'Sub P()',
            'End Sub',
        ].join(eol);
        const tags = scanDocTags(leadingDocLines(source, source.indexOf('Sub P')))!;
        expect(tags.map(tag => tag.tag)).toEqual(['summary', 'param', 'remarks', 'param']);
        expect(tags[0].open).toEqual({start:source.indexOf('<summary>'),end:source.indexOf('<summary>')+9});
        expect(tags[0].end).toBe(source.indexOf('</summary>')+10);
        expect(tags[1].name).toBe('A&B');
        expect(source.slice(tags[1].nameSpan!.start,tags[1].nameSpan!.end)).toBe('A&amp;B');
        expect(tags[1].open.end).toBe(source.indexOf('>Value')+1);
        expect(tags[1].end).toBe(source.indexOf('</param>')+8);
        expect(tags[2].end).toBe(source.indexOf('<remarks />')+11);
        expect(tags[3].name).toBe('Last');
        expect(source.slice(tags[3].nameSpan!.start,tags[3].nameSpan!.end)).toBe('Last');
    });

    it('maps every tag in a large documentation block to its original name', () => {
        const source = Array.from({length:2000},(_,i)=>"''' <param name=\"P"+i+"\" />").join('\n')+'\nSub P()\nEnd Sub';
        const tags=scanDocTags(leadingDocLines(source,source.indexOf('Sub P')))!;
        expect(tags).toHaveLength(2000);
        tags.forEach((tag,i)=>expect(source.slice(tag.nameSpan!.start,tag.nameSpan!.end)).toBe('P'+i));
    });
});
