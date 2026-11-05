// Issue #1622. Diagnostics knew the members of a referenced project's exposed
// class, but completion and hover did not: the editor's member context left
// the referenced surfaces out. ProbeLibrary is the oracle's provider, measured
// in Excel, Word, PowerPoint and Access 16.0: Widget is exposed, Solo exposed
// with a default instance, Thing private.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('../helpers/vscodeMock')).vscodeMock({
	workspace: {
		onDidCloseTextDocument: vi.fn(() => ({ dispose: () => undefined })),
	},
}));

import * as vscode from 'vscode';
import { resolveMemberCompletionNamed, resolveMemberCompletions } from '../../src/analyzer/completion/memberAccess';
import { hostObjectModelForTokens } from '../../src/analyzer/host/hostRegistry';
import type { ProjectEngine } from '../../src/projectEngine';
import { readModules } from '../../src/vba/projectService';
import { toMemberCompletionContext, VbaEditorProjectContextService, type EditorProjectContext } from '../../src/vbaEditorProjectContext';
import { VbaProjectIndexService } from '../../src/vbaProjectIndexService';
import { VbaSymbolIndex } from '../../src/vbaSymbolIndex';
import { oracleProjectReferences, type OracleFixtureModule } from '../helpers/oracleModuleInput';

interface OracleCase { id: string; projectReference?: { name: string; modules: OracleFixtureModule[] } }
const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const providerCase = corpus.cases.find((oracleCase) => oracleCase.id === 'referenced_class_enum_compile')!;
const surface = oracleProjectReferences(providerCase as Parameters<typeof oracleProjectReferences>[0]).referencedProjectSurfaces![0];

const HOSTS = [['excel', 'Excel'], ['word', 'Word'], ['powerpoint', 'PowerPoint'], ['access', 'Access']] as const;

function contextFor(host: typeof HOSTS[number][0], hostName: string): EditorProjectContext {
	return {
		moduleName: 'Module1', moduleKind: 'standard', host,
		hostModel: hostObjectModelForTokens([host]),
		referencedProjectSurfaces: [surface],
		referencedLookupOrder: { libraries: ['stdole', 'Office', 'ProbeLibrary'], hostName },
	};
}

/** The member names offered with the caret after the last `.` of `line`. */
function offered(context: EditorProjectContext, declarations: string, line: string): string[] {
	const source = `Sub Main()\r\n${declarations}\r\n    ${line}`;
	return resolveMemberCompletions(source, source.length, toMemberCompletionContext(context)).map((item) => item.name).sort();
}

describe('a referenced class\'s members in completion and hover (issue #1622)', () => {
	it.each(HOSTS)('offers an exposed class\'s Public members after a variable of it, in %s', (host, hostName) => {
		const context = contextFor(host, hostName);
		const members = offered(context, '    Dim w As ProbeLibrary.Widget', 'w.');
		expect(members).toEqual(['Child', 'Poke', 'Size', 'Twice', 'Value']);
		// No Private, Friend or event member.
		for (const absent of ['mSecret', 'Inner', 'Changed']) { expect(members).not.toContain(absent); }
		// A bare type name reaches the class too.
		expect(offered(context, '    Dim w As Widget', 'w.')).toEqual(members);
		// A chain: Child is a Widget.
		expect(offered(context, '    Dim w As ProbeLibrary.Widget', 'w.Child.')).toEqual(members);
		// A class with a default instance offers its members by name.
		expect(offered(context, '', 'Solo.')).toEqual(['Value']);
	});

	it('offers nothing for a private class, and nothing without the referenced project read', () => {
		const context = contextFor('excel', 'Excel');
		expect(offered(context, '    Dim t As ProbeLibrary.Thing', 't.')).toEqual([]);
		const unread: EditorProjectContext = { moduleName: 'Module1', moduleKind: 'standard' };
		expect(offered(unread, '    Dim w As ProbeLibrary.Widget', 'w.')).toEqual([]);
	});

	it('resolves one member of such a class, as hover asks for it', () => {
		const context = contextFor('word', 'Word');
		const source = 'Sub Main()\r\n    Dim w As ProbeLibrary.Widget\r\n    x = w.Value';
		const member = resolveMemberCompletionNamed(source, source.length - 'Value'.length, 'Value', toMemberCompletionContext(context));
		expect(member?.name).toBe('Value');
	});

	it('carries the referenced projects the project index read into the editor context', async () => {
		// The real consumer workbook through the real project index, with its
		// reference read as the oracle's ProbeLibrary: the checkout the
		// fixtures were made in holds a library without a class.
		const consumer = path.join(__dirname, '..', 'fixtures', 'binaries', 'ProbeConsumerFixture.xlsm');
		const bridge = {
			call: vi.fn(async (method: string, payload: { path: string }) => {
				if (method !== 'readModules') { throw new Error(`Unexpected bridge call ${method}`); }
				const entries = readModules(payload.path);
				entries[0].referencedProjects = [surface];
				return entries;
			}),
		} as unknown as ProjectEngine;
		const contexts = new VbaEditorProjectContextService(new VbaProjectIndexService(new VbaSymbolIndex(bridge)));
		const source = 'Option Explicit\r\nPublic Sub Main()\r\n    Dim w As ProbeLibrary.Widget\r\n    w.';
		const value = `xlide-vba:/${consumer.replace(/\\/g, '/').replace(/^\//, '')}/Module1.bas`;
		const document = { uri: { scheme: 'xlide-vba', path: value.slice('xlide-vba:'.length), toString: () => value }, version: 1, getText: () => source };
		(vscode.workspace.textDocuments as unknown[]).push(document);
		const context = await contexts.buildEditorProjectContextWithin(document as unknown as vscode.TextDocument, source, 30_000);
		expect(context?.referencedProjectSurfaces?.map((candidate) => candidate.name)).toEqual(['ProbeLibrary']);
		expect(context?.referencedLookupOrder).toEqual({ libraries: ['stdole', 'Office', 'ProbeLibrary'], hostName: 'Excel' });
		const names = resolveMemberCompletions(source, source.length, toMemberCompletionContext(context!)).map((item) => item.name).sort();
		expect(names).toEqual(['Child', 'Poke', 'Size', 'Twice', 'Value']);
	});

	it('leaves the member context as it was when no referenced project is read', () => {
		const member = toMemberCompletionContext({ moduleName: 'Module1', referencedProjectSurfaces: [] });
		expect(member.referencedProjectSurfaces).toBeUndefined();
		expect(member.referencedLookupOrder).toBeUndefined();
	});
});
