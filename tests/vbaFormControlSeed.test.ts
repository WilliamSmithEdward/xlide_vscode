import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';

// Issue #48. A form's controls are declared by its DESIGNER, not its text, so
// the host seeds them. That seed has three states, and only two were honoured:
// a list, a vouched-for-EMPTY list, and NO ANSWER because the designer could not
// be read. Reading no answer as an empty one claimed every control the form's
// own code-behind names was undeclared.
//
// It reaches a developer the most ordinary way there is: the VBE stops handing
// out a UserForm's designer once the form has been shown, so running your own
// form turned its code red against source that had compiled a minute earlier.

const BEHIND = [
	'Option Explicit',
	'',
	'Private Sub Recalculate()',
	'    lblPreview.Caption = txtAmount.Text',
	'End Sub',
	'',
].join('\n');

function undeclared(implicitMembers: readonly { name: string; type: string }[] | undefined): string[] {
	return analyzeModule(BEHIND, {
		moduleName: 'frmLoan',
		moduleKind: 'userform',
		knownIdentifiers: new Set<string>(),
		...(implicitMembers === undefined ? {} : { implicitMembers }),
	})
		.filter((diagnostic) => diagnostic.code === 'undeclared-variable')
		.map((diagnostic) => (diagnostic.message.match(/'([^']+)'/) ?? [])[1] ?? '?')
		.sort();
}

describe("a form's control list: absent is not empty", () => {
	it('says nothing when the designer was read and named the controls', () => {
		expect(undeclared([
			{ name: 'txtAmount', type: 'MSForms.TextBox' },
			{ name: 'lblPreview', type: 'MSForms.Label' },
		])).toEqual([]);
	});

	it('reports them when the designer was read and the form has NONE', () => {
		// The case worth protecting: vouched-for-empty is a real claim, so a name
		// that looks like a control really is undeclared.
		expect(undeclared([])).toEqual(['lblPreview', 'txtAmount']);
	});

	it('says nothing when the designer could not be read at all', () => {
		// No control list, no absence claim - the same line the member rule draws.
		expect(undeclared(undefined)).toEqual([]);
	});

	it('still reports an undeclared name outside a form', () => {
		// The suppression is scoped to forms; nothing else seeds implicit members,
		// and a standard module has no designer to be missing.
		const src = 'Option Explicit\nPublic Sub T()\n    NoSuchName = 1\nEnd Sub\n';
		const found = analyzeModule(src, {
			moduleName: 'Module1',
			moduleKind: 'standard',
			knownIdentifiers: new Set<string>(),
		}).filter((diagnostic) => diagnostic.code === 'undeclared-variable');
		expect(found).toHaveLength(1);
	});
});
