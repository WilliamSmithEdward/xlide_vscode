import { describe, expect, it } from 'vitest';
import { collectVbaColors, vbaColorPresentations } from '../src/vbaColors';

describe('VBA color literals', () => {
	it('does not mistake comparisons in calls for color assignments', () => {
		expect(collectVbaColors('Debug.Print Me.BackColor = 255\nCall Check(Me.BackColor = 255)')).toEqual([]);
	});
	it('uses signed VBA hex semantics and accepts explicit Long colors', () => {
		const colors = collectVbaColors('Me.BackColor = &HFFFF\nMe.ForeColor = &HFFFF&');
		expect(colors).toHaveLength(1);
		expect(colors[0]).toMatchObject({ red: 255, green: 255, blue: 0 });
	});
	it('recognizes numbered statements and assignments in single-line If branches', () => {
		expect(collectVbaColors('10 Me.BackColor = 255\nIf ok Then Me.BackColor = 0 Else Me.BackColor = 255')).toHaveLength(3);
	});
	it('does not paint shadowing enum members as built-in colors', () => {
		expect(collectVbaColors('Enum Custom\nvbRed = 123\nEnd Enum\nSub T()\nx = vbRed\nEnd Sub')).toEqual([]);
	});
	it('reads RGB channels and replaces the whole expression', () => {
		const source = 'ActiveCell.Interior.Color = RGB(255, 128, 0)';
		const [color] = collectVbaColors(source);
		expect(color).toMatchObject({ red: 255, green: 128, blue: 0 });
		expect(source.slice(color.start, color.end)).toBe('RGB(255, 128, 0)');
	});
	it('reads packed decimal and hexadecimal colors in VBA byte order', () => {
		const colors = collectVbaColors('ActiveCell.Interior.Color = &HFF&\nMe.BackColor = 16711680');
		expect(colors.map(c => [c.red, c.green, c.blue])).toEqual([[255, 0, 0], [0, 0, 255]]);
	});
	it('reads named and qualified VBA colors', () => {
		const source = 'x = vbRed\ny = ColorConstants.vbBlue\nz = VBA.RGB(0, 255, 0)';
		expect(collectVbaColors(source).map(c => source.slice(c.start, c.end))).toEqual(['vbRed', 'ColorConstants.vbBlue', 'VBA.RGB(0, 255, 0)']);
	});
	it('ignores comments, strings, unrelated numbers, dynamic expressions and system colors', () => {
		const source = `' RGB(255, 0, 0)\nRem vbRed\nx = "vbBlue"\nx = 255\nMe.BackColor = &H8000000F&\nMe.BackColor = 255 + x\nx = RGB(r, 0, 0)\nx = RGB(256, 0, 0)\nx = obj.vbRed`;
		expect(collectVbaColors(source)).toEqual([]);
	});
	it('supports continuation lines and colon-separated assignments', () => {
		expect(collectVbaColors('Me.BackColor = _\n &HFF&: Me.ForeColor = 0')).toHaveLength(2);
	});
	it('does not treat source bindings with built-in color names as literals', () => {
		expect(collectVbaColors('Dim vbRed As Long\nSub T()\nx = vbRed\nEnd Sub')).toEqual([]);
		expect(collectVbaColors('Sub T(ByVal vbRed As Long)\nx = vbRed\nEnd Sub')).toEqual([]);
	});
	it('replaces the full library-qualified constant', () => {
		const source = 'x = VBA.ColorConstants.vbRed';
		const [color] = collectVbaColors(source);
		expect(source.slice(color.start, color.end)).toBe('VBA.ColorConstants.vbRed');
	});
	it('emits valid RGB, packed hexadecimal and decimal presentations', () => {
		expect(vbaColorPresentations(1, 0.5, 0)).toEqual(['RGB(255, 128, 0)', '&H0080FF&', '33023']);
	});
});
