// The folder layout's arithmetic: where a module sits once its `@Folder`
// annotation is read (github.com/WilliamSmithEdward/xlide_vscode/issues/66).

import { describe, expect, it } from 'vitest';
import { buildFolderTree, folderPathChain, type FolderTreeFolder } from '../src/folderTree';

interface Mod { name: string; folder?: string }

const mod = (name: string, folder?: string): Mod => (folder ? { name, folder } : { name });

/** The shape as a reader sees it: folder names nested, modules named inline. */
function outline<T extends { name: string }>(folders: FolderTreeFolder<T>[], modules: T[]): unknown[] {
	return [
		...folders.map((f) => [f.name, outline(f.folders, f.modules)]),
		...modules.map((m) => m.name),
	];
}

describe('building the tree', () => {
	it('nests by the dotted path and keeps unannotated modules at the root', () => {
		const tree = buildFolderTree([
			mod('Ledger', 'Accounts.Ledger'),
			mod('Posting', 'Accounts.Ledger'),
			mod('Invoice', 'Accounts.Billing'),
			mod('ThisWorkbook'),
		]);
		expect(outline(tree.folders, tree.modules)).toEqual([
			['Accounts', [
				['Billing', ['Invoice']],
				['Ledger', ['Ledger', 'Posting']],
			]],
			'ThisWorkbook',
		]);
	});

	it('puts folders first and sorts them by name without regard to case', () => {
		const tree = buildFolderTree([
			mod('Zed', 'zeta'),
			mod('Alpha', 'Alpha'),
			mod('Loose'),
			mod('Mid', 'middle'),
		]);
		expect(tree.folders.map((f) => f.name)).toEqual(['Alpha', 'middle', 'zeta']);
		expect(tree.modules.map((m) => m.name)).toEqual(['Loose']);
	});

	it('keeps the module order it was given, which is the flat tree order', () => {
		const tree = buildFolderTree([
			mod('Sheet1', 'Shared'),
			mod('Helpers', 'Shared'),
			mod('Tools', 'Shared'),
		]);
		expect(tree.folders[0].modules.map((m) => m.name)).toEqual(['Sheet1', 'Helpers', 'Tools']);
	});
});

describe('folders that differ only in case', () => {
	it('merges them, and keeps the first spelling seen', () => {
		const tree = buildFolderTree([
			mod('A', 'Accounts.Ledger'),
			mod('B', 'ACCOUNTS.ledger'),
			mod('C', 'accounts'),
		]);
		expect(tree.folders).toHaveLength(1);
		expect(tree.folders[0].name).toBe('Accounts');
		expect(tree.folders[0].folders[0].name).toBe('Ledger');
		expect(tree.folders[0].folders[0].modules.map((m) => m.name)).toEqual(['A', 'B']);
		expect(tree.folders[0].modules.map((m) => m.name)).toEqual(['C']);
	});

	it('names the merged folder the same way at every depth, for the tooltip', () => {
		const tree = buildFolderTree([mod('A', 'Accounts.Ledger'), mod('B', 'accounts.LEDGER')]);
		expect(tree.folders[0].folders[0].path).toBe('Accounts.Ledger');
	});
});

describe('what a collapsed folder can say about itself', () => {
	it('counts the modules under it, however deep', () => {
		const tree = buildFolderTree([
			mod('A', 'Accounts'),
			mod('B', 'Accounts.Ledger'),
			mod('C', 'Accounts.Ledger.Old'),
			mod('D'),
		]);
		expect(tree.folders[0].moduleCount).toBe(3);
		expect(tree.folders[0].folders[0].moduleCount).toBe(2);
	});
});

describe('an annotation that names nothing usable', () => {
	it('leaves the module at the root rather than making an empty folder', () => {
		const tree = buildFolderTree([mod('A', ''), mod('B')]);
		expect(tree.folders).toEqual([]);
		expect(tree.modules.map((m) => m.name)).toEqual(['A', 'B']);
	});
});

describe('the chain the tree opens on the way to a module', () => {
	it('names every folder from the outermost in', () => {
		expect(folderPathChain('Accounts.Billing.Reminders'))
			.toEqual(['Accounts', 'Accounts.Billing', 'Accounts.Billing.Reminders']);
	});

	it('is empty for a module at the root', () => {
		expect(folderPathChain(undefined)).toEqual([]);
		expect(folderPathChain('')).toEqual([]);
	});
});
