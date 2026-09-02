import { describe, expect, it, vi } from 'vitest';
import type * as VscodeType from 'vscode';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import {
	applyOpenDocumentSources,
	openModuleSourceForProject,
	openModuleSourceMapForProject,
	openXlideModuleSources,
	type VbaOpenDocumentLike,
} from '../src/vbaOpenDocuments';

function doc(path: string, source: string): VbaOpenDocumentLike {
	return {
		uri: {
			scheme: 'xlide-vba',
			path,
			toString: () => path,
		} as VscodeType.Uri,
		getText: () => source,
	};
}

describe('vbaOpenDocuments', () => {
	it('collects only decodable XLIDE project module documents', () => {
		const project = path.join(path.sep, 'one', 'book.xlsm');
		const docs = [
			doc('/one/book.xlsm/Module1.bas', 'Sub One()\nEnd Sub\n'),
			{
				uri: { scheme: 'file', path: '/one/Module1.bas', toString: () => 'file' } as VscodeType.Uri,
				getText: () => 'ignored',
			},
			doc('/one/not-a-workbook.txt/Module1.bas', 'ignored'),
		];

		expect(openXlideModuleSources(docs)).toEqual([
			{
				projectPath: project,
				moduleName: 'Module1',
				source: 'Sub One()\nEnd Sub\n',
			},
		]);
	});

	it('finds an open source only within the requested project', () => {
		const oneWorkbook = path.join(path.sep, 'one', 'book.xlsm');
		const missingWorkbook = path.join(path.sep, 'missing', 'book.xlsm');
		const docs = [
			doc('/one/book.xlsm/Module1.bas', 'Sub FromOne()\nEnd Sub\n'),
			doc('/two/book.xlsm/Module1.bas', 'Sub FromTwo()\nEnd Sub\n'),
		];

		expect(openModuleSourceForProject(oneWorkbook, 'module1', docs))
			.toContain('FromOne');
		expect(openModuleSourceForProject(missingWorkbook, 'Module1', docs))
			.toBeUndefined();
	});

	it('builds a same-project open source map keyed by module identity', () => {
		const oneWorkbook = path.join(path.sep, 'one', 'book.xlsm');
		const docs = [
			doc('/one/book.xlsm/Module1.bas', 'Sub FromOne()\nEnd Sub\n'),
			doc('/one/book.xlsm/MODULE2.bas', 'Sub FromTwo()\nEnd Sub\n'),
			doc('/two/book.xlsm/Module1.bas', 'Sub OtherWorkbook()\nEnd Sub\n'),
		];

		const map = openModuleSourceMapForProject(oneWorkbook, docs);

		expect(map.get('module1')).toContain('FromOne');
		expect(map.get('module2')).toContain('FromTwo');
		expect(map.size).toBe(2);
	});

	it('overlays same-project modules without mutating cached modules', () => {
		const oneWorkbook = path.join(path.sep, 'one', 'book.xlsm');
		const modules = [
			{ moduleName: 'Module1', source: 'Sub SavedOne()\nEnd Sub\n', type: 'standard' },
			{ moduleName: 'Module2', source: 'Sub SavedTwo()\nEnd Sub\n', type: 'standard' },
		];
		const docs = [
			doc('/one/book.xlsm/module1.bas', 'Sub LiveOne()\nEnd Sub\n'),
			doc('/two/book.xlsm/Module2.bas', 'Sub OtherWorkbook()\nEnd Sub\n'),
		];

		const overlaid = applyOpenDocumentSources(modules, oneWorkbook, docs);

		expect(overlaid[0].source).toContain('LiveOne');
		expect(overlaid[1].source).toContain('SavedTwo');
		expect(modules[0].source).toContain('SavedOne');
		expect(overlaid[0]).not.toBe(modules[0]);
	});
});
