import { beforeEach, describe, expect, it } from 'vitest';
import {
	clearXlideWriteAudit,
	formatChangeSummary,
	formatChangeSummaryDetails,
	recentXlideWriteAudits,
	recordXlideWriteAudit,
} from '../src/xlideWriteAudit';

describe('XLIDE write audit', () => {
	beforeEach(() => {
		clearXlideWriteAudit();
	});

	it('records recent write audit facts without command arguments', () => {
		recordXlideWriteAudit({
			timestamp: '2026-06-01T12:00:00.000Z',
			command: 'xlide.writeModule',
			operation: 'write-module',
			outcome: 'succeeded',
			projectPath: 'C:\\Work\\Book.xlsm',
			moduleName: 'Module1',
			summary: 'Write module: 1 changed',
		});

		expect(recentXlideWriteAudits()).toEqual([
			{
				timestamp: '2026-06-01T12:00:00.000Z',
				command: 'xlide.writeModule',
				operation: 'write-module',
				outcome: 'succeeded',
				projectPath: 'C:\\Work\\Book.xlsm',
				moduleName: 'Module1',
				summary: 'Write module: 1 changed',
			},
		]);
	});

	it('caps audit entries', () => {
		for (let i = 0; i < 105; i++) {
			recordXlideWriteAudit({
				timestamp: `2026-06-01T12:00:${String(i).padStart(2, '0')}.000Z`,
				command: 'xlide.writeModule',
				operation: 'write-module',
				outcome: 'succeeded',
				moduleName: `Module${i}`,
				summary: `Write module: ${i}`,
			});
		}

		const recent = recentXlideWriteAudits(200);
		expect(recent).toHaveLength(100);
		expect(recent[0].moduleName).toBe('Module5');
		expect(recent[99].moduleName).toBe('Module104');
	});

	it('formats compact and detailed change summaries', () => {
		const summary = {
			operation: 'Export modules',
			changed: ['Module1.bas'],
			skipped: ['NewModule.bas'],
			removed: ['Stale.bas'],
			failed: ['Bad.bas'],
		};

		expect(formatChangeSummary(summary)).toBe(
			'Export modules: 1 changed, 1 skipped, 1 removed, 1 failed',
		);
		expect(formatChangeSummaryDetails(summary)).toEqual([
			'Export modules: 1 changed, 1 skipped, 1 removed, 1 failed',
			'Changed:',
			'  - Module1.bas',
			'Skipped:',
			'  - NewModule.bas',
			'Removed:',
			'  - Stale.bas',
			'Failed:',
			'  - Bad.bas',
		]);
	});
});
