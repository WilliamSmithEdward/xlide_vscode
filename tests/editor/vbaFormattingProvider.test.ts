import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('../helpers/vscodeMock')).vscodeMock({
	workspace: {
		onDidCloseTextDocument: vi.fn(() => ({ dispose: () => undefined })),
	},
}));

import type * as vscode from 'vscode';
import { conditionalDeclarationScanCountForTests } from '../../src/analyzer/format/conditionalNameCase';
import { VbaFormattingProvider } from '../../src/vbaFormattingProvider';
import { VbaProjectIndexService } from '../../src/vbaProjectIndexService';
import { VbaSymbolIndex } from '../../src/vbaSymbolIndex';
import { encodeModuleUri } from '../../src/xlideFileSystem';
import { fakeProjectEngine, type FakeBridgeModule } from '../helpers/fakeProjectEngine';

// The project path as the module URI decodes it, which on Windows has backslashes.
const BOOK = process.platform === 'win32' ? 'C:\\work\\Book.xlsm' : '/work/Book.xlsm';

function providerFor(modules: FakeBridgeModule[]): { provider: VbaFormattingProvider; engineCalls: () => number } {
	const engine = fakeProjectEngine({ [BOOK]: modules });
	return {
		provider: new VbaFormattingProvider(new VbaProjectIndexService(new VbaSymbolIndex(engine)), () => undefined),
		engineCalls: () => vi.mocked(engine.call).mock.calls.length,
	};
}

/** Formats `text` as the open editor on `uri`, and returns the result. */
async function format(provider: VbaFormattingProvider, uri: vscode.Uri, text: string): Promise<string> {
	const lines = text.split('\n');
	const document = {
		uri,
		languageId: 'vba',
		getText: () => text,
		positionAt: () => ({ line: 0, character: 0 }),
	} as unknown as vscode.TextDocument;
	const edits = await provider.provideDocumentFormattingEdits(
		document,
		{ tabSize: 4, insertSpaces: true },
		{ isCancellationRequested: false } as vscode.CancellationToken,
	) as vscode.TextEdit[];
	for (const edit of edits) {
		// The provider builds `new Range(line, 0, line, end)`; the mock's Range
		// keeps the first argument, the line, as `start`.
		lines[edit.range.start as unknown as number] = edit.newText;
	}
	return lines.join('\n');
}

describe('VbaFormattingProvider - conditional-compilation names', () => {
	const conditionalConstants = 'ProjArg = 1';
	const modules = (): FakeBridgeModule[] => [
		{ name: 'First', type: 'standard', source: '#Const DebugMode = 1\n#Const SharedFlag = 1\n', projectConditionalConstants: conditionalConstants },
		{ name: 'Second', type: 'standard', source: '#Const SHAREDFLAG = 1\nPublic mac As Long\n', projectConditionalConstants: conditionalConstants },
	];

	it("spells them from the project's arguments, every module's #Const in project order, and the open text", async () => {
		const { provider } = providerFor(modules());
		// The open First declares DEBUGMODE, not the DebugMode saved.
		const open = '#Const DEBUGMODE = 1\n#Const SharedFlag = 1\n#If debugmode Or sharedflag Or projarg Or vba7 Then\n#End If';
		expect(await format(provider, encodeModuleUri(BOOK, 'First'), open)).toBe(
			'#Const DEBUGMODE = 1\n#Const SHAREDFLAG = 1\n#If DEBUGMODE Or SHAREDFLAG Or ProjArg Or VBA7 Then\n#End If',
		);
	});

	it("lets another module's ordinary declaration spell a compiler constant", async () => {
		const { provider } = providerFor(modules());
		expect(await format(provider, encodeModuleUri(BOOK, 'First'), '#If Mac Or WIN64 Then\n#End If')).toBe(
			'#If mac Or Win64 Then\n#End If',
		);
	});

	it('reads a module the project does not hold yet from its open text, after the others', async () => {
		const { provider } = providerFor(modules());
		expect(await format(provider, encodeModuleUri(BOOK, 'Third'), '#Const SharedFlag = 1\n#If sharedflag Then\n#End If')).toBe(
			'#Const SharedFlag = 1\n#If SharedFlag Then\n#End If',
		);
	});

	it('spells them from the file alone outside a project, without loading one', async () => {
		const { provider, engineCalls } = providerFor(modules());
		const loose = { scheme: 'untitled', path: 'Untitled-1', toString: () => 'untitled:Untitled-1' } as unknown as vscode.Uri;
		expect(await format(provider, loose, '#Const Loose = 1\n#If LOOSE Or win64 Or sharedflag Then\n#End If')).toBe(
			'#Const Loose = 1\n#If Loose Or Win64 Or sharedflag Then\n#End If',
		);
		expect(engineCalls()).toBe(0);
	});

	it('reads each unchanged module once across formats', async () => {
		const { provider } = providerFor(modules());
		const open = '#Const DebugMode = 1\n#If debugmode Then\n#End If';
		const before = conditionalDeclarationScanCountForTests();
		await format(provider, encodeModuleUri(BOOK, 'First'), open);
		expect(conditionalDeclarationScanCountForTests() - before).toBe(2);
		await format(provider, encodeModuleUri(BOOK, 'First'), open);
		expect(conditionalDeclarationScanCountForTests() - before).toBe(2);
		// A changed open text is read again; the other module is not.
		await format(provider, encodeModuleUri(BOOK, 'First'), `${open}\n#Const Other = 1`);
		expect(conditionalDeclarationScanCountForTests() - before).toBe(3);
	});
});
