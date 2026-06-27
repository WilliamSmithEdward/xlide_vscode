import { beforeEach, describe, expect, it } from 'vitest';
import {
	clearXlideCommandLog,
	errorCategoryForSupportLog,
	recentXlideCommands,
	recordXlideCommand,
} from '../src/xlideCommandLog';

describe('XLIDE command log', () => {
	beforeEach(() => {
		clearXlideCommandLog();
	});

	it('keeps command ids and outcomes without command arguments', () => {
		recordXlideCommand({
			timestamp: '2026-06-01T12:00:00.000Z',
			command: 'xlide.analyzeCurrentModule',
			outcome: 'started',
		});
		recordXlideCommand({
			timestamp: '2026-06-01T12:00:01.000Z',
			command: 'xlide.analyzeCurrentModule',
			outcome: 'succeeded',
			durationMs: 10,
		});

		expect(recentXlideCommands()).toEqual([
			{
				timestamp: '2026-06-01T12:00:00.000Z',
				command: 'xlide.analyzeCurrentModule',
				outcome: 'started',
			},
			{
				timestamp: '2026-06-01T12:00:01.000Z',
				command: 'xlide.analyzeCurrentModule',
				outcome: 'succeeded',
				durationMs: 10,
			},
		]);
	});

	it('caps stored command log entries', () => {
		for (let i = 0; i < 105; i++) {
			recordXlideCommand({
				timestamp: `2026-06-01T12:00:${String(i).padStart(2, '0')}.000Z`,
				command: `xlide.command${i}`,
				outcome: 'succeeded',
			});
		}

		const recent = recentXlideCommands(200);
		expect(recent).toHaveLength(100);
		expect(recent[0].command).toBe('xlide.command5');
		expect(recent[99].command).toBe('xlide.command104');
	});

	it('classifies common support error categories without preserving messages', () => {
		expect(errorCategoryForSupportLog(new Error('PermissionError: WinError 32'))).toBe(
			'workbook-locked',
		);
		// Read-only Excel opens surface as WinError 5 / Access is denied rather than
		// WinError 32; these must still classify (and report) as workbook-locked.
		expect(
			errorCategoryForSupportLog(new Error("[WinError 5] Access is denied: 'C:\\\\book.xlsm'")),
		).toBe('workbook-locked');
		expect(errorCategoryForSupportLog(new Error('[Errno 13] EACCES: permission'))).toBe(
			'workbook-locked',
		);
		expect(errorCategoryForSupportLog(new Error('python backend exited unexpectedly'))).toBe(
			'python-backend',
		);
		expect(errorCategoryForSupportLog(new Error('User cancelled operation'))).toBe('cancelled');
		expect(errorCategoryForSupportLog(new Error('something else'))).toBe('unknown');
	});
});
