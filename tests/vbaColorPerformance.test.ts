import { describe, expect, it, vi } from 'vitest';
import { collectVbaColors } from '../src/vbaColors';
import * as symbols from '../src/analyzer/symbols/buildModuleSymbols';
import * as cursor from '../src/analyzer/completion/cursorContext';

describe('color scan work bounds', () => {
	it('scans thousands of numeric colors without per-color prefix or symbol analysis', () => {
		const prefix = vi.spyOn(cursor, 'completionCursorContext');
		const build = vi.spyOn(symbols, 'buildModuleSymbols');
		try {
			const source = `Sub ColorPerf()\n${'Me.BackColor = 255\n'.repeat(5000)}End Sub`;
			expect(collectVbaColors(source)).toHaveLength(5000);
			expect(prefix).not.toHaveBeenCalled();
			expect(build).not.toHaveBeenCalled();
		} finally { prefix.mockRestore(); build.mockRestore(); }
	});
	it('builds scope bindings once for many named colors and reuses the source snapshot', () => {
		const build = vi.spyOn(symbols, 'buildModuleSymbols');
		try {
			const source = Array.from({ length: 1000 }, (_, i) => `Sub ColorScope${i}()\nDim unused${i} As Long\nx = vbRed\nEnd Sub`).join('\n');
			expect(collectVbaColors(source)).toHaveLength(1000);
			expect(collectVbaColors(source)).toHaveLength(1000);
			expect(build).toHaveBeenCalledTimes(1);
		} finally { build.mockRestore(); }
	});
	it('stops a scan when cancelled without returning partial decorations', () => {
		let probes = 0;
		expect(collectVbaColors('Me.BackColor = 255\n'.repeat(1000), () => ++probes > 2)).toEqual([]);
		expect(probes).toBe(3);
	});
});
