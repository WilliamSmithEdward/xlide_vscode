import { describe, expect, it } from 'vitest';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { powerpointReferenceData } from '../../src/analyzer/host/powerpointObjectModelData';

describe('host-specific reference examples', () => {
    it('corrects Word chart receivers in the shipped PowerPoint examples', () => {
        const examples = Object.values(getPowerPointObjectModel().types).map(type => type.doc?.example).filter(Boolean) as string[];
        const curated = Object.values(powerpointReferenceData().types).filter(type => type.doc?.example?.startsWith("' Requires a chart as shape 1"));
        expect(curated).toHaveLength(36);
        expect(examples.some(example => /ActiveDocument\.InlineShapes/.test(example))).toBe(false);
        expect(getPowerPointObjectModel().types['PowerPoint.ChartData'].doc?.example).toContain('With ActivePresentation.Slides(1).Shapes(1).Chart.ChartData');
    });

    it.each(['Excel', 'Word', 'Access', undefined])('preserves examples for %s', async host => {
        const { typeDoc } = await import('../../scripts/reference-curation.mjs');
        const example = 'With ActiveDocument.InlineShapes(1)\r\n    MsgBox .HasChart\r\nEnd With';
        expect(typeDoc({ example }, 300, host)?.example).toBe(example);
    });

    it('changes only the known opening receiver and preserves the chart operation', async () => {
        const { typeDoc } = await import('../../scripts/reference-curation.mjs');
        const example = 'With ActiveDocument.InlineShapes(1).Chart\n    .ChartTitle.Text = "ActiveDocument.InlineShapes(1)"\nEnd With';
        expect(typeDoc({ example }, 300, 'PowerPoint')?.example).toBe(
            "' Requires a chart as shape 1 on slide 1 of the active presentation.\nWith ActivePresentation.Slides(1).Shapes(1).Chart\n    .ChartTitle.Text = \"ActiveDocument.InlineShapes(1)\"\nEnd With");
        expect(typeDoc({ example: 'MsgBox "ActiveDocument.InlineShapes(1)"' }, 300, 'PowerPoint')?.example)
            .toBe('MsgBox "ActiveDocument.InlineShapes(1)"');
    });
});
