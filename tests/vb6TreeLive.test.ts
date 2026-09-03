import { describe, expect, it, vi } from 'vitest';
import * as path from 'path';
import { readFileSync } from 'fs';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	EventEmitter: class {
		event = vi.fn();
		fire = vi.fn();
	},
	window: { showErrorMessage: vi.fn() },
	workspace: {
		findFiles: vi.fn().mockResolvedValue([]),
		workspaceFolders: [{ uri: { fsPath: 'C:\\work' } }],
	},
}));

import { ProjectExplorer } from '../src/projectExplorer';
import * as service from '../src/vba/projectService';

// The tree over a real VB6 project on disk, driven through the same bridge
// calls the extension makes. The engine's own listSubs is proven elsewhere;
// what this pins is that the tree ASKS for the procedures of every module
// kind a VB6 project has and turns them into rows.

const PROJECT = path.join(__dirname, 'fixtures', 'vb6', 'RunAsTrustedInstaller', 'Project1.vbp');

/** A bridge that answers from the real engine, as the extension host does. */
function engineBridge(): { call: <T>(method: string, params: Record<string, unknown>) => Promise<T> } {
	return {
		call: async <T>(method: string, params: Record<string, unknown>): Promise<T> => {
			const filePath = params.path as string;
			switch (method) {
				case 'listModules':
					return service.listModules(filePath) as T;
				case 'listSubs':
					return service.listSubs(filePath, params.module as string) as T;
				case 'getProtectionInfo':
					return service.getProtectionInfo(filePath) as T;
				default:
					throw new Error(`unexpected bridge call: ${method}`);
			}
		},
	};
}

describe('the tree over a VB6 project on disk', () => {
	it('lists every module, and the procedures under each', async () => {
		const explorer = new ProjectExplorer(engineBridge());
		const project = { kind: 'project' as const, label: path.basename(PROJECT), filePath: PROJECT };
		const modules = await explorer.getChildren(project);
		expect(modules.length).toBeGreaterThan(0);
		expect(modules.map((m) => m.kind)).toEqual(modules.map(() => 'module'));

		// Every module the project has must offer its procedures, and a module
		// node must be expandable for the user to ever see them.
		let procedures = 0;
		for (const module of modules) {
			expect(explorer.getTreeItem(module).collapsibleState, module.label).not.toBe(0);
			const children = await explorer.getChildren(module);
			expect(children.some((c) => c.kind === 'loadError'), `${module.label} failed to load`).toBe(false);
			const subs = children.filter((c) => c.kind === 'sub');
			procedures += subs.length;
			for (const sub of subs) {
				expect(sub.label).toMatch(/^(Sub|Function|Property (Get|Let|Set)) \w+$/);
				expect(sub.moduleName).toBe(module.moduleName);
				expect(sub.line).toBeGreaterThan(0);
				expect(explorer.getTreeItem(sub).command?.command).toBeTruthy();
			}
		}
		expect(procedures).toBeGreaterThan(0);

		// A form's rows: the designer first, then its procedures.
		const form = modules.find((m) => m.moduleType === 'userform');
		expect(form).toBeDefined();
		const formChildren = await explorer.getChildren(form!);
		expect(formChildren[0].kind).toBe('designer');
		expect(formChildren.filter((c) => c.kind === 'sub').length).toBeGreaterThan(0);
	});

	it('expands the module the active editor is showing, file or virtual document', () => {
		// Opening a module marks it active, and an active module renders
		// Expanded - which is how clicking a module in the tree reveals its
		// procedures. That tracking read the module out of an `xlide-vba:`
		// URI, so a VB6 module, whose document is its own file, was never the
		// active one and its row never opened.
		const activation = readFileSync(path.join(__dirname, '..', 'src', 'extension.ts'), 'utf8');
		expect(activation).toContain('moduleLocationOfDocument(editor.document)');
		expect(activation).not.toContain('pending = decodeModuleUri(editor.document.uri);');

		const explorer = new ProjectExplorer(engineBridge());
		const node = {
			kind: 'module' as const, label: 'Form1', filePath: PROJECT,
			moduleName: 'Form1', moduleType: 'userform',
		};
		expect(explorer.getTreeItem(node).collapsibleState).toBe(1); // Collapsed
		explorer.setActiveModule(PROJECT, 'Form1');
		expect(explorer.getTreeItem(node).collapsibleState).toBe(2); // Expanded
	});
});
