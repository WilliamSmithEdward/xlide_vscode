import { describe, expect, it } from 'vitest';
import { getHostMembers } from '../src/analyzer/host/hostModel';
import { hostObjectModelForToken } from '../src/analyzer/host/hostRegistry';
import { applicationMemberNames } from '../src/analyzer/diagnostics/analysisContext';
import { analyzeModule, resolveMemberCompletions } from '../src/analyzer';
import { byCode } from './helpers/diagnostics';
import { analyzeProjectModule } from './diagnostics/helpers';

// Issue #56: the model carries the type library's plumbing, and until now
// nothing marked it. A member is now flagged when the library says hidden or
// restricted, or when its name is one VBA cannot write at all - and a flagged
// member is still KNOWN (resolved, hovered, coloured) while never being
// OFFERED, because accepting the proposal would not compile.
describe('hidden host members', () => {
	it('flags the dispatch plumbing the VBE will not let you type', () => {
		const members = getHostMembers('Excel.Workbook');
		const codeName = members.find((m) => m.name === '_CodeName');
		expect(codeName, 'the model must still CARRY _CodeName').toBeDefined();
		expect(codeName?.hidden).toBe(true);
	});

	it('leaves the everyday members alone', () => {
		// Range.Rows / Columns / EntireRow are marked FNONBROWSABLE in the
		// library because they are duplicated on Application and Global.
		// Treating that flag as "hidden" would delete them from completion.
		const range = getHostMembers('Excel.Range');
		for (const name of ['Rows', 'Columns', 'EntireRow', 'EntireColumn', 'Value']) {
			const member = range.find((m) => m.name === name);
			expect(member, `Range.${name} must exist`).toBeDefined();
			expect(member?.hidden, `Range.${name} must not be hidden`).toBeFalsy();
		}
	});

	it('does not OFFER a hidden member, while still resolving it', () => {
		const source = 'Sub T()\r\n    ThisWorkbook.\r\nEnd Sub\r\n';
		const offered = resolveMemberCompletions(source, source.indexOf('ThisWorkbook.') + 'ThisWorkbook.'.length);
		expect(offered.length).toBeGreaterThan(20);
		expect(offered.some((c) => c.name === 'Worksheets')).toBe(true);
		expect(offered.some((c) => c.name.startsWith('_'))).toBe(false);
		expect(offered.some((c) => c.hidden)).toBe(false);
	});

	describe('members the library hides and the documentation leaves out', () => {
		// Found in ReDim's HostProbe, which compiles:
		//     Debug.Print ThisWorkbook.Names.Count, ThisWorkbook.FullName, ThisWorkbook.Title
		// reported member-not-found on Title. The model came from the reference
		// documentation, which does not list a member the type library marks
		// hidden, so every closed surface - a document module, Worksheet, Chart -
		// refused them. Measured on EXCEL.EXE: Workbook lacked 28, Worksheet 10,
		// Chart 52, and no visible member was missing anywhere.
		const memberNotFound = (body: string, modules: Parameters<typeof analyzeProjectModule>[1] = []) => {
			const src = `Option Explicit\nPublic Sub T()\n${body}End Sub\n`;
			return byCode(analyzeProjectModule(src, modules, 'Probe'), 'member-not-found');
		};

		it('carries them, marked hidden', () => {
			const hidden = (type: string, name: string) => getHostMembers(`Excel.${type}`).find((m) => m.name === name);
			for (const [type, name] of [['Workbook', 'Title'], ['Workbook', 'Author'], ['Worksheet', 'OnEntry'], ['Worksheet', 'DisplayAutomaticPageBreaks']]) {
				expect(hidden(type, name), `Excel.${type}.${name}`).toMatchObject({ hidden: true });
			}
		});

		it('accepts one on a document module, as in the report', () => {
			const document = { moduleName: 'ThisWorkbook', moduleKind: 'document' as const, documentType: 'workbook' as const, source: '' };
			expect(memberNotFound('    Debug.Print ThisWorkbook.Title\n', [document])).toHaveLength(0);
		});

		it('accepts one on Worksheet, which is closed', () => {
			expect(memberNotFound('    Dim ws As Worksheet\n    ws.OnEntry = "Handler"\n')).toHaveLength(0);
		});

		it('still refuses a name the library does not have at all', () => {
			expect(memberNotFound('    Dim ws As Worksheet\n    ws.OnEntryy = "Handler"\n')).toHaveLength(1);
		});

		it('never offers one', () => {
			const source = 'Sub T()\r\n    Dim wb As Workbook\r\n    wb.\r\nEnd Sub\r\n';
			const offered = resolveMemberCompletions(source, source.indexOf('wb.') + 'wb.'.length).map((c) => c.name);
			expect(offered).toContain('FullName');
			expect(offered).not.toContain('Title');
		});

		it('does not make a hidden Application member callable bare', () => {
			// `Save` is a hidden method of _Application and no member of _Global,
			// so `Application.Save` compiles and a bare `Save` does not. VBA
			// resolves bare names through Global, which answers for itself.
			expect(getHostMembers('Excel.Application').find((m) => m.name === 'Save')).toMatchObject({ hidden: true });
			// Unknown bare calls are only judged against a whole project.
			const src = 'Option Explicit\nPublic Sub T()\n    Application.Save\n    Save\nEnd Sub\n';
			const unknown = byCode(analyzeProjectModule(src, [], 'Probe'), 'unknown-call');
			expect(unknown.map((d) => src.slice(d.span.start, d.span.end))).toEqual(['Save']);
		});
	});

	describe('in Word and PowerPoint', () => {
		// Their dumps list no hidden member at all: measured against MSWORD.OLB
		// and MSPPT.OLB, Word's model lacked 332 and PowerPoint's 200, and not
		// one visible member. Those models are never exhaustive, so a member
		// they lack is never reported missing - but a hidden GLOBAL is a bare
		// name, and hiding the Office Assistant, a staple of older macros, read
		// as "Variable not defined".
		const errors = (host: 'word' | 'powerpoint', body: string) => {
			const src = `Option Explicit\nPublic Sub T()\n    ${body}\nEnd Sub\n`;
			return analyzeProjectModule(src, [], 'Probe', { host })
				.filter((d) => d.severity === 'error')
				.map((d) => `${d.code}: ${src.slice(d.span.start, d.span.end)}`);
		};

		it('accepts a hidden global called bare', () => {
			expect(errors('word', 'Assistant.Visible = False')).toEqual([]);
			expect(errors('word', 'AnswerWizard.ClearFileList')).toEqual([]);
			expect(errors('powerpoint', 'Assistant.Visible = False')).toEqual([]);
			expect(errors('powerpoint', 'Debug.Print TypeName(Dialogs)')).toEqual([]);
		});

		it('still refuses a hidden Application member called bare that Global does not carry', () => {
			// ShowMe is hidden on Word's _Application and absent from _Global.
			expect(errors('word', 'ShowMe')).toEqual(['unknown-call: ShowMe']);
			expect(errors('word', 'Application.ShowMe')).toEqual([]);
		});

		it('carries the rest marked hidden, so they resolve and are never offered', () => {
			const word = hostObjectModelForToken('word')!;
			expect(getHostMembers('Word.Document', word).find((m) => m.name === 'AutoSummarize')).toMatchObject({ hidden: true });
			const ppt = hostObjectModelForToken('powerpoint')!;
			expect(getHostMembers('PowerPoint.Presentation', ppt).find((m) => m.name === 'HasRevisionInfo')).toMatchObject({ hidden: true });
		});
	});

	it('keeps a hidden Application member in bare scope where there is no Global', () => {
		// Access binds Application itself bare - its type library makes it the
		// app object - so hidden or not, its members are callable unqualified.
		// Only where a Global interface answers for bare names do Application's
		// hidden members stay out.
		const members = [
			{ name: 'Visible', kind: 'property' as const },
			{ name: 'SecretThing', kind: 'method' as const, hidden: true },
		];
		const model = (globalType?: string) => ({
			source: 'test',
			types: {
				'Test.Application': { displayName: 'Application', members },
				...(globalType ? { [globalType]: { displayName: 'Global', members: [] } } : {}),
			},
			aliases: { application: 'Test.Application' },
			globals: { Application: 'Test.Application' },
			...(globalType ? { globalType } : {}),
		}) as unknown as Parameters<typeof applicationMemberNames>[0];
		expect([...applicationMemberNames(model())]).toEqual(['visible', 'secretthing']);
		expect([...applicationMemberNames(model('Test.Global'))]).toEqual(['visible']);
	});

	it('offers no member whose name VBA could not write', () => {
		// Whatever the receiver, a proposal must be something the VBE compiles.
		for (const receiver of ['ThisWorkbook.', 'Application.', 'ActiveSheet.']) {
			const source = `Sub T()\r\n    ${receiver}\r\nEnd Sub\r\n`;
			const offered = resolveMemberCompletions(source, source.indexOf(receiver) + receiver.length);
			const unwritable = offered.filter((c) => !/^[A-Za-z]/.test(c.name));
			expect(unwritable.map((c) => c.name), `${receiver} offered an unwritable name`).toEqual([]);
		}
	});
});
