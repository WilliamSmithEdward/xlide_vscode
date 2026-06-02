import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
	clearXlideOutputLog,
	createRecordedOutputChannel,
	recentXlideOutputLog,
	redactSupportLogLine,
} from '../src/xlideOutputLog';

describe('XLIDE output log', () => {
	beforeEach(() => {
		clearXlideOutputLog();
	});

	it('redacts path-like values before support export', () => {
		expect(
			redactSupportLogLine('Starting Python bridge: C:\\Users\\William\\repo\\python.exe'),
		).toBe('Starting Python bridge: <redacted>.exe');
		expect(
			redactSupportLogLine('Open xlide-vba:///C%3A/Users/William/Book.xlsm?module=Module1'),
		).toBe('Open <redacted-uri>');
	});

	it('records appendLine calls through the wrapped output channel', () => {
		const raw = {
			name: 'XLIDE',
			append: vi.fn(),
			appendLine: vi.fn(),
			replace: vi.fn(),
			clear: vi.fn(),
			show: vi.fn(),
			hide: vi.fn(),
			dispose: vi.fn(),
		};
		const out = createRecordedOutputChannel(raw);

		out.appendLine('Reading C:\\Users\\William\\Documents\\Book.xlsm');
		out.append('partial write is not captured');

		expect(raw.appendLine).toHaveBeenCalledWith('Reading C:\\Users\\William\\Documents\\Book.xlsm');
		expect(raw.append).toHaveBeenCalledWith('partial write is not captured');
		expect(recentXlideOutputLog()).toEqual([
			{
				timestamp: expect.any(String),
				line: 'Reading <redacted>.xlsm',
			},
		]);
	});

	it('caps recent output entries', () => {
		for (let i = 0; i < 255; i++) {
			const out = createRecordedOutputChannel({
				name: 'XLIDE',
				append: vi.fn(),
				appendLine: vi.fn(),
				replace: vi.fn(),
				clear: vi.fn(),
				show: vi.fn(),
				hide: vi.fn(),
				dispose: vi.fn(),
			});
			out.appendLine(`line ${i}`);
		}

		const recent = recentXlideOutputLog(300);
		expect(recent).toHaveLength(250);
		expect(recent[0].line).toBe('line 5');
		expect(recent[249].line).toBe('line 254');
	});
});
