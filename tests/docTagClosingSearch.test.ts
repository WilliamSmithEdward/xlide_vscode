import { describe, expect, it, vi } from 'vitest';
import { leadingDocLines, scanDocTags } from '../src/analyzer/docs/docComment';

function scan(source: string) {
	return scanDocTags(leadingDocLines(source, source.indexOf('Sub P')))!;
}

describe('documentation closing tag searches', () => {
	it.each([false, true])('bounds repeated suffix searches when a final close exists: %s', (closed) => {
		const count = 1000;
		const source = Array.from({length: count}, (_, i) => "''' <param name=\"P" + i + '\">text')
			.join('\n') + (closed ? '</param>' : '') + '\nSub P()\nEnd Sub';
		const original = String.prototype.indexOf;
		let searched = 0;
		const spy = vi.spyOn(String.prototype, 'indexOf').mockImplementation(function (this: string, search: string, position?: number) {
			if (search === '</param>') {
				searched += this.length - (position ?? 0);
			}
			return original.call(this, search, position);
		});
		let tags: ReturnType<typeof scan>;
		try {
			tags = scan(source);
		} finally {
			spy.mockRestore();
		}
		expect(searched).toBeLessThan(source.length * 2);
		expect(tags!).toHaveLength(count);
		expect(tags!.slice(0, -1).every(tag => tag.end === undefined)).toBe(true);
		expect(tags!.at(-1)!.end !== undefined).toBe(closed);
	});

	it('advances past an old close while keeping tag types independent', () => {
		const source = [
			"''' <param name=\"A\">one</param>",
			"''' <summary>start",
			"''' <param name=\"B\">unfinished",
			"''' <param name=\"C\" />",
			"''' <param name=\"D\">four</param></summary>",
			"''' <param name=\"E\">never closed",
			'Sub P()', 'End Sub',
		].join('\r\n');
		const tags = scan(source);
		expect(tags.map(tag => tag.text)).toEqual(['one', 'start <param name="B">unfinished <param name="C" /> <param name="D">four</param>', undefined, '', 'four', undefined]);
		expect(tags[4].end).toBe(source.indexOf('four</param>') + 'four</param>'.length);
	});

	it('preserves reopening markers inside another tag attribute', () => {
		const tags = scan("''' <param name=\"A\">one <remarks x=\"<param?\">text</remarks> </param>\nSub P()\nEnd Sub");
		expect(tags).toHaveLength(2);
		expect(tags[0].end).toBeUndefined();
	});
});
