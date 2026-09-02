import { describe, expect, it } from 'vitest';
import { parseTwinSource, resolveSurfaces } from '../scripts/twinbasic-vb-surface.mjs';

// The reader of twinBASIC's VB package source (roadmap_vb6_support.md,
// Slice 4). The package is the oracle's own statement of which VB6 members
// it implements; misreading it would flag the model wrongly in both
// directions, so the shapes the package uses are pinned here on a synthetic
// sample: stacked and inline attributes, [Unimplemented], #If guards, an
// Inherits chain, Get/Let merging, a CoClass over a default interface with
// an Extends chain, bodies that must not leak locals, and bare interface
// declarations.

const SAMPLE = [
	'#If FEATURE_BUTTON Then',
	'[ClassId("1")]',
	'Private Class ButtonBaseCtl',
	'    Inherits BaseControlRect',
	'    [WithDispatchForwarding] Implements Control',
	'    #Region "STATE"',
	'        [CustomDesigner("x")]',
	'        [Description("The text on the face")]',
	'            Public Caption As String = ""',
	'        [Unimplemented]',
	'            Public RightToLeft As Boolean',
	'        Public ReadOnly ControlType As Long',
	'        Protected Internal As Long',
	'        Public Const kMax As Long = 3',
	'    #End Region',
	'    Public Property Get Value() As Boolean',
	'        Dim local As Long',
	'        Public NotAMember As String',
	'        Value = True',
	'    End Property',
	'    Public Property Let Value(ByVal rhs As Boolean)',
	'    End Property',
	'    [Hidden]',
	'    Public Property Get Secret() As Long',
	'    End Property',
	'    [DefaultDesignerEvent]',
	'    [Description("")]',
	'        Event Click()',
	'    Event MouseDown(Button As Integer, Shift As Integer)',
	'    #If FEATURE_HELP Then',
	'    Public HelpContextID As Long',
	'    Public Sub ShowWhatsThis()',
	'    End Sub',
	'    #Else',
	'    Public NeverHere As Long',
	'    #End If',
	'    Private Sub Hidden()',
	'    End Sub',
	'End Class',
	'',
	'Class CommandButton',
	'    Inherits ButtonBaseCtl',
	'End Class',
	'#End If',
	'',
	'Private Class BaseControlRect',
	'    Inherits BaseControl',
	'    Public Property Get Left() As Single',
	'    End Property',
	'    Public Property Let Left(ByVal rhs As Single)',
	'    End Property',
	'End Class',
	'',
	'Private Class BaseControl',
	'    Public Tag As String',
	'    Public Property Get Name() As String',
	'    End Property',
	'End Class',
	'',
	'Private Interface _AppBase Extends stdole.IUnknown',
	'    Property Get Major() As Integer',
	'End Interface',
	'',
	'Private Interface _App Extends _AppBase',
	'    Property Get Path() As String',
	'    [Unimplemented]',
	'    Property Get LogMode() As Long',
	'    Sub StartLogging(ByVal LogTarget As String, ByVal LogMode As Long)',
	'End Interface',
	'',
	'[CoClassId("2")]',
	'Public CoClass App',
	'    [Default] Interface _App',
	'End CoClass',
	'',
	'Public Interface Control',
	'End Interface',
	'',
].join('\r\n');

function surfaceOf(name: string) {
	const surface = resolveSurfaces(parseTwinSource(SAMPLE, 'Sample.twin')).get(name);
	expect(surface, name).toBeDefined();
	return surface!;
}

function member(name: string, memberName: string, kind = '') {
	return surfaceOf(name).members.find((m: { name: string; kind: string }) =>
		m.name.toLowerCase() === memberName.toLowerCase() && (!kind || m.kind === kind));
}

describe('reading twinBASIC package source', () => {
	it('walks the Inherits chain and keeps only public members', () => {
		const names = surfaceOf('CommandButton').members.map((m: { name: string }) => m.name);
		expect(names).toEqual(expect.arrayContaining(['Caption', 'RightToLeft', 'Value', 'Left', 'Tag', 'Name', 'Click', 'MouseDown', 'HelpContextID', 'ShowWhatsThis', 'kMax']));
		expect(names).not.toContain('Internal');
		expect(names).not.toContain('Hidden');
		expect(names).not.toContain('NotAMember');
		expect(names).not.toContain('local');
		expect(names).not.toContain('NeverHere');
	});

	it('reads the attributes that matter and the feature guard', () => {
		expect(member('CommandButton', 'RightToLeft')?.unimplemented).toBe(true);
		expect(member('CommandButton', 'Caption')?.unimplemented).toBe(false);
		expect(member('CommandButton', 'Caption')?.description).toBe('The text on the face');
		expect(member('CommandButton', 'Secret')?.hidden).toBe(true);
		expect(member('CommandButton', 'HelpContextID')?.feature).toBe('FEATURE_BUTTON And FEATURE_HELP');
		expect(member('CommandButton', 'Caption')?.feature).toBe('FEATURE_BUTTON');
	});

	it('merges Get and Let into one property with an access label', () => {
		expect(member('CommandButton', 'Value', 'property')?.access).toBe('read/write');
		expect(member('CommandButton', 'Secret', 'property')?.access).toBe('read-only');
		// `Public ReadOnly X As T` is a read-only property named X, not a field named ReadOnly.
		expect(member('CommandButton', 'ControlType', 'property')?.access).toBe('read-only');
		expect(member('CommandButton', 'ReadOnly')).toBeUndefined();
		expect(member('CommandButton', 'Caption', 'property')?.access).toBe('read/write');
		expect(member('CommandButton', 'Left', 'property')?.declaredType).toBe('Single');
		expect(surfaceOf('CommandButton').members.filter((m: { name: string }) => m.name === 'Value')).toHaveLength(1);
	});

	it('records events with their parameters and the forwarded interface', () => {
		expect(member('CommandButton', 'MouseDown', 'event')?.params).toBe('Button As Integer, Shift As Integer');
		expect(member('CommandButton', 'Click', 'event')?.params).toBe('');
		// The public class is a thin Inherits over the base that implements.
		expect(surfaceOf('CommandButton').implements).toEqual([{ name: 'Control', forwarding: true }]);
		expect(surfaceOf('ButtonBaseCtl').implements).toEqual([{ name: 'Control', forwarding: true }]);
		expect(surfaceOf('CommandButton').bases).toEqual(['ButtonBaseCtl']);
	});

	it('gives a CoClass the members of its default interface, Extends chain included', () => {
		const app = surfaceOf('App');
		expect(app.kind).toBe('coclass');
		expect(app.members.map((m: { name: string }) => m.name)).toEqual(expect.arrayContaining(['Path', 'LogMode', 'StartLogging', 'Major']));
		expect(member('App', 'LogMode')?.unimplemented).toBe(true);
		expect(member('App', 'StartLogging')?.kind).toBe('method');
		expect(member('App', 'StartLogging')?.params).toBe('ByVal LogTarget As String, ByVal LogMode As Long');
		expect(member('App', 'Path')?.access).toBe('read-only');
	});

	it('stays private where the package says so', () => {
		const surfaces = resolveSurfaces(parseTwinSource(SAMPLE));
		expect(surfaces.get('ButtonBaseCtl')?.access).toBe('private');
		expect(surfaces.get('CommandButton')?.access).toBe('public');
		expect(surfaces.get('Control')?.members).toEqual([]);
	});
});
