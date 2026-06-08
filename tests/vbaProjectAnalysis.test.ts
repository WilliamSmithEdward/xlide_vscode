import { describe, expect, it, vi } from 'vitest';
import { buildVbaProjectIndexAsync, type VbaProjectModuleInput } from '../src/vbaProjectAnalysis';

function modules(count: number): VbaProjectModuleInput[] {
	return Array.from({ length: count }, (_, index) => ({
		moduleName: `Module${index + 1}`,
		moduleKind: 'standard',
		source: `Public Sub Test${index + 1}()\nEnd Sub\n`,
	}));
}

describe('buildVbaProjectIndexAsync', () => {
	it('yields cooperatively between module chunks', async () => {
		vi.useFakeTimers();
		try {
			let settled = false;
			const promise = buildVbaProjectIndexAsync(modules(3), undefined, {
				yieldEveryModules: 1,
			}).then((project) => {
				settled = true;
				return project;
			});

			await Promise.resolve();
			expect(settled).toBe(false);

			await vi.runOnlyPendingTimersAsync();
			expect(settled).toBe(false);

			await vi.runAllTimersAsync();
			const project = await promise;
			expect(settled).toBe(true);
			expect(project.visibleProcedureNames('Module3').has('test3')).toBe(true);
		} finally {
			vi.useRealTimers();
		}
	});

	it('honors cooperative cancellation between chunks', async () => {
		vi.useFakeTimers();
		try {
			const cancelled = new Error('cancelled');
			let checks = 0;
			const promise = buildVbaProjectIndexAsync(modules(3), undefined, {
				yieldEveryModules: 1,
				cancelIfRequested: () => {
					checks += 1;
					if (checks > 1) {
						throw cancelled;
					}
				},
			});
			const rejection = expect(promise).rejects.toBe(cancelled);

			await vi.runOnlyPendingTimersAsync();
			await rejection;
		} finally {
			vi.useRealTimers();
		}
	});
});
