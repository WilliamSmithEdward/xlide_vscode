import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// A VB6 form or UserControl IS the class its designer makes it. That class's
// members are the module's own: `Me.Arrange` reaches the MDI form's, and a
// bare `PropertyChanged` binds to the UserControl's, though the module's text
// declares neither. Without that the project analyzer called both undefined,
// while live diagnostics stayed quiet - the same code red on one surface and
// clean on the other.

const MDI_SOURCE = [
	'Option Explicit',
	'',
	'Private Sub mnuWindowTile_Click()',
	'    Me.Arrange vbTileHorizontal',
	'End Sub',
	'',
].join('\r\n');

const CONTROL_SOURCE = [
	'Option Explicit',
	'',
	'Public Property Let Caption(ByVal newValue As String)',
	'    PropertyChanged "Caption"',
	'    Me.Refresh',
	'End Property',
	'',
].join('\r\n');

function analyze(source: string, moduleName: string, extra: Record<string, unknown>) {
	return analyzeVbaModuleSource({
		source,
		moduleName,
		host: 'vb6',
		// The cross-module rules only run when the caller supplies these.
		knownProcedures: new Set<string>(),
		knownIdentifiers: new Set<string>(),
		...extra,
	} as Parameters<typeof analyzeVbaModuleSource>[0]).diagnostics.map((d) => `${d.code}`);
}

describe('a module gets the members of the class its designer makes it', () => {
	it('reaches an MDI form s own methods through Me', () => {
		const withClass = analyze(MDI_SOURCE, 'mdiMain', { moduleType: 'userform', moduleKind: 'userform', designerClass: 'VB.MDIForm' });
		expect(withClass).not.toContain('member-not-found');
		// The negative has no counterpart here: `Me`'s surface only becomes
		// exhaustive enough to report a missing member once the cross-module
		// project context is in play, which single-module analysis has not
		// got. That end of it is covered where the whole project is analyzed.
	});

	it('binds a bare call to a UserControl s own method', () => {
		const withClass = analyze(CONTROL_SOURCE, 'ctlBadge', { moduleType: 'usercontrol', moduleKind: 'class', designerClass: 'VB.UserControl' });
		expect(withClass).not.toContain('unknown-call');
		expect(withClass).not.toContain('member-not-found');
		const without = analyze(CONTROL_SOURCE, 'ctlBadge', { moduleType: 'usercontrol', moduleKind: 'class' });
		expect(without).toContain('unknown-call');
	});

	it('still reports a call that no designer class defines', () => {
		const source = [
			'Option Explicit',
			'',
			'Private Sub Go()',
			'    NoSuchMethod "x"',
			'End Sub',
			'',
		].join('\r\n');
		const codes = analyze(source, 'ctlBadge', { moduleType: 'usercontrol', moduleKind: 'class', designerClass: 'VB.UserControl' });
		expect(codes).toContain('unknown-call');
	});

	it('asserts nothing for a class the model does not carry', () => {
		const codes = analyze(CONTROL_SOURCE, 'ctlBadge', { moduleType: 'usercontrol', moduleKind: 'class', designerClass: 'Made.Up' });
		expect(codes).toContain('unknown-call');
	});
});
