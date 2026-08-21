import { describe, expect, it } from 'vitest';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import { getExcelObjectModel, type HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import {
	getHostMembers,
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
	resolveMemberReturnType,
} from '../src/analyzer/host/hostModel';
import {
	analyzeModule,
	resolveIdentifierCompletions,
	resolveMemberCompletions,
	resolveSignatureHelp,
} from '../src/analyzer';

// Issue #25. The measured number this replaces: Word's ThisDocument answered
// 1 member - the Sub in its own module - because the only model that existed
// was Excel's and the honest alternative was silence.

const MODELS: ReadonlyArray<[string, () => HostObjectModel]> = [
	['word', getWordObjectModel],
	['powerpoint', getPowerPointObjectModel],
	['access', getAccessObjectModel],
];

// Origin labels are built from hostName (issue #28), so a misspelled name
// would ship in every hover and completion detail of that host.
it('every model names its application for origin labels', () => {
	expect(getExcelObjectModel().hostName).toBe('Excel');
	expect(getWordObjectModel().hostName).toBe('Word');
	expect(getPowerPointObjectModel().hostName).toBe('PowerPoint');
	expect(getAccessObjectModel().hostName).toBe('Access');
});

// Every Office VBA project auto-references the shared Office library, so
// msoTrue and friends are everyday legal code in every host. Excel merged
// them from the start; the generated hosts must too, or Option Explicit
// flags msoTrue as an undeclared variable in a Word module.
describe('shared Office constants reach every host model', () => {
	it('resolves msoTrue in Word, PowerPoint, and Access', () => {
		for (const getModel of [getWordObjectModel, getPowerPointObjectModel, getAccessObjectModel]) {
			const constant = resolveHostConstant('msoTrue', getModel());
			expect(constant?.value).toBe(-1);
			expect(constant?.type).toBe('MsoTriState');
		}
	});

	it('keeps the host library first on shared chart-enum names', () => {
		// Word and PowerPoint re-export the Xl chart enums the Office library
		// also carries (299 same-value names, measured); either source must
		// answer with the shared value.
		for (const getModel of [getWordObjectModel, getPowerPointObjectModel]) {
			const constant = resolveHostConstant('xlAbove', getModel());
			expect(constant?.value).toBe(0);
			expect(constant?.type).toBe('XlConstants');
		}
	});

	it('keeps msoTrue out of the undeclared-variable rule in every host', () => {
		const src = 'Option Explicit\nSub T()\n    Dim n As Long\n    n = msoTrue\nEnd Sub\n';
		for (const host of ['word', 'powerpoint', 'access', 'excel']) {
			const findings = analyzeModule(src, { host, knownIdentifiers: new Set<string>() })
				.map((d) => d.message);
			expect(findings, host).toEqual([]);
		}
		// The rule itself still works: an invented name stays a finding.
		const bogus = 'Option Explicit\nSub T()\n    Dim n As Long\n    n = msoNotAConstant\nEnd Sub\n';
		const flagged = analyzeModule(bogus, { host: 'word', knownIdentifiers: new Set<string>() });
		expect(flagged.some((d) => d.message.includes('msoNotAConstant'))).toBe(true);
	});
});

describe.each(MODELS)('the %s object model', (_host, getModel) => {
	it('declares every global against a type it actually carries', () => {
		const model = getModel();
		for (const [name, qualified] of Object.entries(model.globals)) {
			expect(model.types[qualified] ?? model.aliases[qualified.toLowerCase()], `${name} -> ${qualified}`)
				.toBeDefined();
		}
	});

	it('proves nothing absent: no type claims exhaustiveness', () => {
		const model = getModel();
		for (const [qualified, type] of Object.entries(model.types)) {
			expect(type.exhaustive ?? false, qualified).toBe(false);
		}
	});

	it('carries documentation, not bare signatures', () => {
		const model = getModel();
		const documented = Object.values(model.types)
			.flatMap((type) => type.members)
			.filter((member) => member.doc?.summary).length;
		expect(documented).toBeGreaterThan(500);
	});
});

describe('Word answers as Word', () => {
	it('ThisDocument is a Document with the everyday members', () => {
		const model = getWordObjectModel();
		const documentType = resolveHostGlobal('ThisDocument', model);
		expect(documentType).toBe('Word.Document');
		const names = getHostMembers(documentType!, model).map((member) => member.name);
		for (const expected of ['Content', 'Paragraphs', 'SaveAs2', 'Range', 'Tables']) {
			expect(names, expected).toContain(expected);
		}
		// The issue's measured "1" becomes the real surface.
		expect(names.length).toBeGreaterThan(200);
	});

	it('does not leak Excel: no Cells on a Document, no Worksheets global', () => {
		const model = getWordObjectModel();
		const names = getHostMembers('Word.Document', model).map((member) => member.name);
		expect(names).not.toContain('Cells');
		expect(resolveHostGlobal('ThisWorkbook', model)).toBeUndefined();
	});

	it('Selection. completes through the member resolver', () => {
		const source = 'Sub T()\r\n    Selection.\r\nEnd Sub\r\n';
		const members = resolveMemberCompletions(source, source.indexOf('Selection.') + 10, {
			model: getWordObjectModel(),
		} as never).map((member) => member.name);
		expect(members).toContain('TypeText');
		expect(members).toContain('Collapse');
	});

	it('knows wd* constants with their values', () => {
		const constant = resolveHostConstant('wdMainTextStory', getWordObjectModel());
		expect(constant?.value).toBe(1);
		expect(constant?.type).toBe('WdStoryType');
	});

	it('offers the Global interface members callable bare (issue #34)', () => {
		// InchesToPoints is in half the PageSetup lines ever written; it is a
		// member of Word's Global interface, not an injected object name.
		const src = 'Sub T()\n    Inches\nEnd Sub\n';
		const items = resolveIdentifierCompletions(src, src.indexOf('Inches') + 6, {
			moduleName: 'M',
			model: getWordObjectModel(),
		});
		const method = items.find((item) => item.name === 'InchesToPoints');
		expect(method?.detail).toBe('InchesToPoints(Inches As Single) As Single');

		const bare = 'Sub T()\n    Recent\nEnd Sub\n';
		const property = resolveIdentifierCompletions(bare, bare.indexOf('Recent') + 6, {
			moduleName: 'M',
			model: getWordObjectModel(),
		}).find((item) => item.name === 'RecentFiles');
		expect(property?.detail).toBe('RecentFiles object');
	});

	it('offers a call tip inside a bare Global method call (issue #34)', () => {
		const src = 'Sub T()\n    TopM = InchesToPoints(1)\nEnd Sub\n';
		const info = resolveSignatureHelp(src, src.indexOf('(1)') + 1, {
			model: getWordObjectModel(),
		} as never);
		expect(info?.label).toBe('InchesToPoints(Inches As Single) As Single');
		expect(info?.parameters.map((p) => p.label)).toEqual(['Inches As Single']);
	});

	it('keeps a Global member out of the undeclared rule in its host only (issue #34)', () => {
		const src = 'Option Explicit\nSub T()\n    Dim n As Single\n    n = InchesToPoints(1)\nEnd Sub\n';
		const word = analyzeModule(src, { host: 'word', knownIdentifiers: new Set<string>() })
			.map((d) => d.message);
		expect(word).toEqual([]);
	});
});

describe('PowerPoint and Access answer as themselves', () => {
	it('ActivePresentation reaches Slides', () => {
		const model = getPowerPointObjectModel();
		const type = resolveHostGlobal('ActivePresentation', model);
		expect(type).toBe('PowerPoint.Presentation');
		expect(getHostMembers(type!, model).map((member) => member.name)).toContain('Slides');
	});

	it('DoCmd carries OpenForm, documented', () => {
		const model = getAccessObjectModel();
		const type = resolveHostGlobal('DoCmd', model);
		expect(type).toBe('Access.DoCmd');
		const openForm = getHostMembers(type!, model).find((member) => member.name === 'OpenForm');
		expect(openForm).toBeDefined();
		expect(openForm?.kind).toBe('method');
	});

	it('the hosts do not bleed into each other or into Excel', () => {
		expect(resolveHostGlobal('ActivePresentation', getWordObjectModel())).toBeUndefined();
		expect(resolveHostGlobal('DoCmd', getPowerPointObjectModel())).toBeUndefined();
		expect(resolveHostGlobal('ActiveDocument', getExcelObjectModel())).toBeUndefined();
	});
});

describe('the shared Office library types reach every host', () => {
    // Every Office VBA project references the Office library, and host
    // members return its types: a Shape's TextFrame2 hands back an Office
    // TextRange2. Without them the chain dead-ends at that hop, and
    // TextRange2 was the one documented PowerPoint object the model lacked.
    it('resolves an Office type in every host', () => {
        for (const getModel of [getExcelObjectModel, getWordObjectModel, getPowerPointObjectModel, getAccessObjectModel]) {
            const model = getModel();
            expect(resolveHostAlias('TextRange2', model), String(model.hostName)).toBe('Office.TextRange2');
        }
    });

    it('chains through a host member that returns an Office type', () => {
        const ppt = getPowerPointObjectModel();
        const frame = resolveMemberReturnType('PowerPoint.Shape', 'TextFrame2', ppt);
        expect(frame).toBe('PowerPoint.TextFrame2');
        const range = resolveMemberReturnType(frame!, 'TextRange', ppt);
        expect(range).toBe('Office.TextRange2');
        expect(getHostMembers(range!, ppt).length).toBeGreaterThan(20);
    });

    it('lets the host library win a shared type name', () => {
        expect(resolveHostAlias('Font', getWordObjectModel())).toBe('Word.Font');
        expect(resolveHostAlias('Font', getExcelObjectModel())).toBe('Excel.Font');
    });

    it('proves nothing absent: no Office type claims exhaustiveness', () => {
        const model = getExcelObjectModel();
        for (const [qualified, type] of Object.entries(model.types)) {
            if (qualified.startsWith('Office.')) {
                expect(type.exhaustive ?? false, qualified).toBe(false);
            }
        }
    });
});
