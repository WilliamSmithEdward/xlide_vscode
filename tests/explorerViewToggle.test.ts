import { describe, expect, it, vi } from 'vitest';
import { createExplorerViewSetter } from '../src/explorerViewToggle';

function harness(persist: (view: string) => Promise<unknown>) {
	const applied: string[] = [];
	const logged: string[] = [];
	const warnPersistFailed = vi.fn();
	const setView = createExplorerViewSetter({
		applyView: (view) => applied.push(view),
		persist: (view) => persist(view),
		log: (line) => logged.push(line),
		warnPersistFailed,
	});
	return { applied, logged, warnPersistFailed, setView };
}

describe('createExplorerViewSetter', () => {
	it('switches the view and remembers it', async () => {
		const persisted: string[] = [];
		const h = harness(async (view) => { persisted.push(view); });

		await h.setView('folders');

		expect(h.applied).toEqual(['folders']);
		expect(persisted).toEqual(['folders']);
		expect(h.logged).toEqual([]);
		expect(h.warnPersistFailed).not.toHaveBeenCalled();
	});

	it('still switches the view when the setting cannot be written (#71)', async () => {
		const h = harness(() => Promise.reject(new Error(
			'Unable to write to User Settings because xlide.explorer.view is not a registered configuration.',
		)));

		await h.setView('folders');

		expect(h.applied).toEqual(['folders']);
		expect(h.logged).toHaveLength(1);
		expect(h.logged[0]).toContain('for this window only');
		expect(h.logged[0]).toContain('not a registered configuration');
	});

	it('does not reject when the setting cannot be written', async () => {
		const h = harness(() => Promise.reject(new Error('nope')));

		await expect(h.setView('tree')).resolves.toBeUndefined();
	});

	it('warns once per window, however many times the write fails', async () => {
		const h = harness(() => Promise.reject(new Error('nope')));

		await h.setView('folders');
		await h.setView('tree');
		await h.setView('folders');

		expect(h.applied).toEqual(['folders', 'tree', 'folders']);
		expect(h.logged).toHaveLength(3);
		expect(h.warnPersistFailed).toHaveBeenCalledTimes(1);
	});

	it('applies the view before it writes, so a slow write cannot delay the tree', async () => {
		const order: string[] = [];
		let release: (() => void) | undefined;
		const setView = createExplorerViewSetter({
			applyView: () => order.push('applied'),
			persist: () => new Promise<void>((resolve) => {
				order.push('persist started');
				release = resolve;
			}),
			log: () => order.push('logged'),
			warnPersistFailed: () => order.push('warned'),
		});

		const pending = setView('folders');
		expect(order).toEqual(['applied', 'persist started']);
		release?.();
		await pending;

		expect(order).toEqual(['applied', 'persist started']);
	});
});
