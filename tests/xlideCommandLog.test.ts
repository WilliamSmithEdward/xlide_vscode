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
			'project-locked',
		);
		// Read-only Excel opens surface as WinError 5 / Access is denied rather than
		// WinError 32; these must still classify (and report) as project-locked.
		expect(
			errorCategoryForSupportLog(new Error("[WinError 5] Access is denied: 'C:\\\\book.xlsm'")),
		).toBe('project-locked');
		expect(errorCategoryForSupportLog(new Error('[Errno 13] EACCES: permission'))).toBe(
			'project-locked',
		);
		// What a save actually meets: XLIDE renames a temp file over the
		// container, and Windows refuses with EPERM while the file is open.
		// Measured against Excel, Word, PowerPoint and Access on Office 16.0.
		expect(errorCategoryForSupportLog(new Error(
			"EPERM: operation not permitted, rename 'C:\\work\\.xlide-30332-1789619210070.tmp' -> 'C:\\work\\Book.xlsm'",
		))).toBe('project-locked');
		expect(errorCategoryForSupportLog(new Error('ENOENT: no such file or directory'))).toBe(
			'project-missing',
		);
		expect(errorCategoryForSupportLog(new Error('User cancelled operation'))).toBe('cancelled');
		expect(errorCategoryForSupportLog(new Error('something else'))).toBe('unknown');
	});
});
