import { describe, expect, it } from 'vitest';
import * as path from 'path';
import {
	listModules,
	readModule,
	resetProjectCacheForTests,
} from '../src/vba/projectService';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// VBA exposes EVERY control as a member of the form, however deeply it nests:
// `Me.PickAir` works when PickAir sits inside a Frame, and `Me.Agree` works
// when Agree sits on a MultiPage's Page. A container's children do not live in
// the form's own `f` stream - each gets a storage of its own - so a reader
// that takes two buffers can never see them (#57). The member surface is read
// by the package walker instead, which recurses.
//
// FormFixture.xlsm, which pins the flat reader, has no containers at all, so
// nesting went unexercised. This file uses the nested fixture.

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

/** Top level, then the Frame's children, then the MultiPage's. */
const EXPECTED = [
	{ name: 'NameLabel', type: 'MSForms.Label' },
	{ name: 'NameBox', type: 'MSForms.TextBox' },
	{ name: 'RegionPick', type: 'MSForms.ComboBox' },
	{ name: 'Taxable', type: 'MSForms.CheckBox' },
	{ name: 'Options', type: 'MSForms.Frame' },
	{ name: 'HoldToggle', type: 'MSForms.ToggleButton' },
	{ name: 'Wizard', type: 'MSForms.MultiPage' },
	{ name: 'Views', type: 'MSForms.TabStrip' },
	{ name: 'ViewNote', type: 'MSForms.Label' },
	{ name: 'Amount', type: 'MSForms.ScrollBar' },
	{ name: 'Steps', type: 'MSForms.SpinButton' },
	{ name: 'Badge', type: 'MSForms.Image' },
	{ name: 'OkButton', type: 'MSForms.CommandButton' },
	{ name: 'HistoryList', type: 'MSForms.ListBox' },
	{ name: 'PickGround', type: 'MSForms.OptionButton' },
	{ name: 'PickAir', type: 'MSForms.OptionButton' },
	{ name: 'Page1', type: 'MSForms.Page' },
	{ name: 'Page2', type: 'MSForms.Page' },
	{ name: 'Agree', type: 'MSForms.CheckBox' },
];

function formEntry() {
	resetProjectCacheForTests();
	const entry = listModules(FIXTURE).find((candidate) => candidate.name === 'EntryForm');
	resetProjectCacheForTests();
	return entry;
}

describe('a container\'s children are members of the form', () => {
	it('lists every control, however deeply it nests', () => {
		expect(formEntry()?.implicitMembers).toEqual(EXPECTED);
	});

	it('names each control once, though a container is reached twice', () => {
		// A Frame arrives with the entries of its own surface AND again in the
		// site sweep that catches a container carrying no record. Listing it
		// twice would put a duplicate member on the form.
		const names = (formEntry()?.implicitMembers ?? []).map((c) => c.name.toLowerCase());
		expect(names).toEqual([...new Set(names)]);
	});

	it('carries the MultiPage, its Pages, and what sits on a Page', () => {
		const members = formEntry()?.implicitMembers ?? [];
		const named = (name: string) => members.find((c) => c.name === name);
		expect(named('Wizard')?.type).toBe('MSForms.MultiPage');
		expect(named('Page1')?.type).toBe('MSForms.Page');
		expect(named('Agree')?.type).toBe('MSForms.CheckBox');
		// The MultiPage's own hidden, unnamed TabStrip site is not a member:
		// only the real TabStrip on the form answers to a name.
		expect(members.filter((c) => c.type === 'MSForms.TabStrip'))
			.toEqual([{ name: 'Views', type: 'MSForms.TabStrip' }]);
	});

	it('code touching a nested control is not called undeclared', () => {
		resetProjectCacheForTests();
		const entries = listModules(FIXTURE).map((entry) => ({
			moduleName: entry.name,
			type: entry.type,
			documentType: entry.documentType,
			source: readModule(FIXTURE, entry.name, true).source,
			implicitMembers: entry.implicitMembers,
		}));
		resetProjectCacheForTests();
		const project = buildVbaProjectIndex(entries);
		const form = entries.find((entry) => entry.moduleName === 'EntryForm')!;
		// The exact false-positive class #57 names: every control here lives
		// inside a Frame or on a Page, and the module declares none of them.
		const source = form.source + [
			'Private Sub Freight()',
			'    PickGround.Value = True',
			'    Me.PickAir.Caption = "Air"',
			'    Agree.Value = False',
			'    Wizard.Value = 0',
			'    Me.Page2.Caption = "Two"',
			'End Sub',
			'',
		].join('\r\n');
		const diagnostics = analyzeVbaModuleSource({
			source,
			moduleName: 'EntryForm',
			moduleType: 'userform',
			moduleKind: 'userform',
			...projectAnalysisOptionsForModule(project, 'EntryForm'),
		} as never).diagnostics;

		expect(diagnostics.filter((d) => d.code === 'undeclared-variable')).toEqual([]);
		expect(diagnostics.filter((d) => d.code === 'member-not-found')).toEqual([]);
	});
});
