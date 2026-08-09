import { describe, expect, it } from 'vitest';
import { ProjectIndex, resolveMemberCompletions } from '../src/analyzer';

// Issue #11: a project Type lost its member surface to a same-named class from
// a referenced type library, and an Enum name offered nothing at all.
function dotOffset(source: string, marker: string): number {
	return source.indexOf(marker) + marker.length;
}

function shapesIndex(typeName: string): ProjectIndex {
	const index = new ProjectIndex();
	index.setModule({
		moduleName: 'Shapes',
		moduleKind: 'standard',
		source: [
			'Option Explicit',
			'',
			`Public Type ${typeName}`,
			'    X As Double',
			'    Y As Double',
			'    Label As String',
			'End Type',
			'',
			'Public Enum Corner',
			'    TopLeft',
			'    TopRight',
			'    BottomLeft',
			'    BottomRight',
			'End Enum',
			'',
			'Private Enum HiddenCorner',
			'    Inner',
			'End Enum',
		].join('\n'),
	});
	index.setModule({ moduleName: 'Uses', moduleKind: 'standard', source: '' });
	return index;
}

function membersAt(index: ProjectIndex, source: string, marker: string): string[] {
	return resolveMemberCompletions(source, dotOffset(source, marker), {
		projectClassMembers: index.projectMemberSurfaces('Uses'),
	}).map((member) => member.name);
}

describe('a project type outranks a same-named library class', () => {
	it('offers the project Type fields for a name Excel also owns', () => {
		// Excel's chart Point class was winning, so `p.` offered 37 members
		// (Application, DataLabel, Explosion, ...) with nothing to indicate the
		// name was ambiguous. The Excel object model owns a lot of ordinary
		// nouns - Point, Border, Font, Shape, Style, Name.
		const got = membersAt(shapesIndex('Point'),
			'Public Sub TypeReceiver()\n    Dim p As Point\n    p.\nEnd Sub\n', 'p.');
		expect(got).toEqual(['X', 'Y', 'Label']);
	});

	it('still offers the project Type fields for a name Excel does not own', () => {
		const got = membersAt(shapesIndex('Anchorage'),
			'Public Sub TypeReceiver()\n    Dim p As Anchorage\n    p.\nEnd Sub\n', 'p.');
		expect(got).toEqual(['X', 'Y', 'Label']);
	});

	it('still reaches the host type when the project declares no such name', () => {
		const got = membersAt(shapesIndex('Anchorage'),
			'Public Sub HostReceiver()\n    Dim r As Range\n    r.\nEnd Sub\n', 'r.');
		expect(got).toContain('Value');
		expect(got).toContain('Cells');
	});
});

describe('an Enum name reaches its constants', () => {
	it('offers enum members after the enum name', () => {
		const got = membersAt(shapesIndex('Anchorage'),
			'Public Sub EnumReceiver()\n    Dim c As Corner\n    c = Corner.\nEnd Sub\n', 'Corner.');
		expect(got).toEqual(['TopLeft', 'TopRight', 'BottomLeft', 'BottomRight']);
	});

	it('does not expose a Private Enum outside its module', () => {
		const got = membersAt(shapesIndex('Anchorage'),
			'Public Sub EnumReceiver()\n    Dim c As Long\n    c = HiddenCorner.\nEnd Sub\n', 'HiddenCorner.');
		expect(got).toEqual([]);
	});

	it('treats an enum-typed variable as a value, not an object', () => {
		// An Enum is a Long: making it a member surface must not make
		// `Dim c As Corner` look like an object that needs Set, and `c.`
		// offers nothing because a Long has no members.
		const got = membersAt(shapesIndex('Anchorage'),
			'Public Sub EnumValue()\n    Dim c As Corner\n    c.\nEnd Sub\n', 'c.');
		expect(got).toEqual([]);
	});
});
