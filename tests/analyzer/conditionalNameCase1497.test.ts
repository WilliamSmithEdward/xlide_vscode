// Format Document gives conditional-compilation names the spelling the VBE
// gives them (issue #1497). The expectations are what Excel, Word,
// PowerPoint and Access 16.0 (build 20430) show on opening the same modules
// (tests/office/directiveCasing.office.ts).

import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('../helpers/vscodeMock')).vscodeMock());

import { conditionalConstNames, conditionalNameSpellings, projectArgumentNames } from '../../src/analyzer/format/conditionalNames';
import { VbaFormattingProvider } from '../../src/vbaFormattingProvider';
import type { VbaProjectIndexService } from '../../src/vbaProjectIndexService';

const BOOK = process.platform === 'win32' ? 'C:/Book.xlsm' : '/work/Book.xlsm';
const BOOK_URI_PATH = BOOK.startsWith('/') ? BOOK : `/${BOOK}`;

describe('the names #Const lines declare', () => {
	it('reads every shape a #Const line can take, in order', () => {
		expect(conditionalConstNames([
			'#Const Flag = 1', '# _', 'Const CONT = 1', '#Const [BRK] = 1', '#Const REMV = 1 Rem note', '#Const COLV = 1:',
			'#If False Then', '#Const DEAD = 1', '#End If', 'Sub P()', '#Const INPROC = 1', 'End Sub',
		].join('\r\n'))).toEqual(['Flag', 'CONT', 'BRK', 'REMV', 'COLV', 'DEAD', 'INPROC']);
	});

	it('does not count a #Const after a statement and a colon', () => {
		expect(conditionalConstNames('x = 1: #Const Y = 1\r\n')).toEqual([]);
	});

	it('reads nothing from a module without one', () => {
		expect(conditionalConstNames('Const X = 1\r\n#If VBA7 Then\r\n#End If\r\n')).toEqual([]);
	});

	it('reads the names of the project arguments', () => {
		expect(projectArgumentNames('ARGNAME = 1 : ArgTwo = -1 : = 2')).toEqual(['ARGNAME', 'ArgTwo']);
	});
});

describe('the spelling of a conditional name', () => {
	it('is VBA\'s for its compiler constants, and not for TWINBASIC', () => {
		const spellings = conditionalNameSpellings(undefined, []);
		expect(['vba6', 'vba7', 'win16', 'win32', 'win64', 'mac'].map((name) => spellings.get(name))).toEqual(['VBA6', 'VBA7', 'Win16', 'Win32', 'Win64', 'Mac']);
		expect(spellings.has('twinbasic')).toBe(false);
	});

	it('is the last declaration: the arguments, then each module in load order', () => {
		const spellings = conditionalNameSpellings('ARGNAME = 1 : ARGTWO = 1', [['Flag', 'WIN64'], ['fLaG', 'ArgTwo']]);
		expect(spellings.get('flag')).toBe('fLaG');
		expect(spellings.get('argname')).toBe('ARGNAME');
		expect(spellings.get('argtwo')).toBe('ArgTwo');
		expect(spellings.get('win64')).toBe('WIN64');
	});
});

function document(moduleName: string, text: string): unknown {
	const lines = text.split('\n');
	return {
		uri: { scheme: 'xlide-vba', path: `${BOOK_URI_PATH}/${moduleName}.bas`, toString: () => `xlide-vba:${BOOK_URI_PATH}/${moduleName}.bas` },
		languageId: 'vba',
		version: 1,
		getText: () => text,
		positionAt: (offset: number) => ({ line: text.slice(0, offset).split('\n').length - 1, character: 0 }),
		lineCount: lines.length,
	};
}

function provider(modules: { moduleName: string; source: string }[], conditionalConstants?: string): VbaFormattingProvider {
	const projects = {
		contextForProject: vi.fn(async () => ({
			references: [],
			project: { visibleIdentifierSymbols: () => [{ name: 'mac' }] },
			modules,
			conditionalConstants,
		})),
	};
	return new VbaFormattingProvider(projects as unknown as VbaProjectIndexService, () => undefined);
}

async function formatted(subject: VbaFormattingProvider, moduleName: string, text: string): Promise<string> {
	const edits = await subject.provideDocumentFormattingEdits(
		document(moduleName, text) as never,
		{ tabSize: 4, insertSpaces: true } as never,
		{ isCancellationRequested: false } as never,
	) as { range: { start: number }; newText: string }[];
	const lines = text.split('\n');
	for (const edit of edits) {
		// The vscode mock keeps a range's start line as a number.
		lines[edit.range.start] = edit.newText;
	}
	return lines.join('\n');
}

describe('Format Document on conditional names (issue #1497)', () => {
	it('spells the compiler constants as VBA does, on directive and code lines', async () => {
		const text = ['#If vba7 Then', '#ElseIf vba6 Then', '#End If', '#If win64 Or win32 Or win16 Then', '#End If', '#If twinbasic Then', '#End If', 'Sub Probe()', '    y = win64', 'End Sub'].join('\n');
		expect(await formatted(provider([{ moduleName: 'XlideA', source: text }]), 'XlideA', text)).toBe(
			['#If VBA7 Then', '#ElseIf VBA6 Then', '#End If', '#If Win64 Or Win32 Or Win16 Then', '#End If', '#If twinbasic Then', '#End If', 'Sub Probe()', '    y = Win64', 'End Sub'].join('\n'),
		);
	});

	it('lets a declaration of the same name take the spelling over', async () => {
		const text = ['#Const WIN64 = 0', '#If win64 Or MAC Then', '#End If'].join('\n');
		expect(await formatted(provider([{ moduleName: 'XlideA', source: text }, { moduleName: 'XlideB', source: 'Public mac As Long\n' }]), 'XlideA', text))
			.toBe(['#Const WIN64 = 0', '#If WIN64 Or mac Then', '#End If'].join('\n'));
	});

	it('takes the last #Const in load order, across modules and from the arguments', async () => {
		const a = ['#Const Flag = 1', '#If FLAG Or argname Or argtwo Then', '#End If', 'Sub Q()', '    z = flag', 'End Sub'].join('\n');
		const b = ['#Const fLaG = 2', '#Const ArgTwo = 1'].join('\n');
		const subject = provider([{ moduleName: 'XlideA', source: a }, { moduleName: 'XlideB', source: b }], 'ARGNAME = 1 : ARGTWO = 1');
		expect(await formatted(subject, 'XlideA', a)).toBe(
			['#Const fLaG = 1', '#If fLaG Or ARGNAME Or ArgTwo Then', '#End If', 'Sub Q()', '    z = fLaG', 'End Sub'].join('\n'),
		);
	});

	it('reads the open module as the editor has it', async () => {
		const saved = '#Const Old = 1\n';
		const editor = ['#Const NewName = 1', '#If newname Or old Then', '#End If'].join('\n');
		expect(await formatted(provider([{ moduleName: 'XlideA', source: saved }]), 'XlideA', editor))
			.toBe(['#Const NewName = 1', '#If NewName Or old Then', '#End If'].join('\n'));
	});

	it('reads another module again only once the project index replaces it', async () => {
		const b = { moduleName: 'XlideB', source: '#Const Common = 1\n' };
		const modules = [{ moduleName: 'XlideA', source: '' }, b];
		const subject = provider(modules);
		const text = '#If common Then\n#End If';
		expect(await formatted(subject, 'XlideA', text)).toBe('#If Common Then\n#End If');
		// The same view with other text is not read again: the index hands
		// out a new view whenever a module changes.
		b.source = '#Const COMMON = 1\n';
		expect(await formatted(subject, 'XlideA', text)).toBe('#If Common Then\n#End If');
		modules[1] = { moduleName: 'XlideB', source: '#Const COMMON = 1\n' };
		expect(await formatted(subject, 'XlideA', text)).toBe('#If COMMON Then\n#End If');
	});
});
