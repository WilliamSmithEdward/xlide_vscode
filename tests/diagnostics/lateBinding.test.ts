import { describe, expect, it } from 'vitest';
import { byCode, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CODE = 'late-bound-friend-member';

// A class whose ToastSlot is Friend - the shape from the real-world report.
const UI_CLASS = [
	'Option Explicit',
	'Private mSlot As Long',
	'Friend Property Get ToastSlot() As Long',
	'    ToastSlot = mSlot',
	'End Property',
	'Friend Property Let ToastSlot(ByVal value As Long)',
	'    mSlot = value',
	'End Property',
	'Public Sub Show()',
	'End Sub',
	'',
].join('\n');

function analyze(source: string, extraModules: Array<{ moduleName: string; moduleKind?: 'class'; source: string }> = []) {
	return analyzeProjectModule(source, [
		{ moduleName: 'ReDimUI', moduleKind: 'class', source: UI_CLASS },
		...extraModules,
	], 'Caller');
}

describe('late-bound Friend member access', () => {
	it('flags a Friend member reached through a Collection element', () => {
		// The reported bug: Collection.Item returns Variant, so the element is
		// late bound and cannot see ToastSlot. Compiles clean; dies at run time.
		const src = [
			'Option Explicit',
			'Public Sub Compact(ByVal live As Collection)',
			'    Dim position As Long',
			'    For position = 2 To live.Count',
			'        Debug.Print live.Item(position).ToastSlot',
			'    Next position',
			'End Sub',
			'',
		].join('\n');

		const hits = byCode(analyze(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ToastSlot');
		expect(hits[0].message).toContain('438');
		expect(hits[0].message).toContain("Friend member of class 'ReDimUI'");
	});

	it('flags the default-member subscript form of the same access', () => {
		const src = [
			'Option Explicit',
			'Public Sub Compact(ByVal live As Collection)',
			'    Debug.Print live(1).ToastSlot',
			'End Sub',
			'',
		].join('\n');

		const hits = byCode(analyze(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ToastSlot');
	});

	it('flags Variant and Object locals, and an untyped Dim', () => {
		for (const declaration of ['Dim item As Variant', 'Dim item As Object', 'Dim item']) {
			const src = [
				'Option Explicit',
				'Public Sub T()',
				`    ${declaration}`,
				'    Debug.Print item.ToastSlot',
				'End Sub',
				'',
			].join('\n');

			const hits = byCode(analyze(src), CODE);
			expect(hits, declaration).toHaveLength(1);
			expect(spanText(src, hits[0]), declaration).toBe('ToastSlot');
		}
	});

	it('stays silent once the element is assigned to a typed local', () => {
		// The documented fix: the typed intermediate makes the call early bound.
		const src = [
			'Option Explicit',
			'Public Sub Compact(ByVal live As Collection)',
			'    Dim candidate As ReDimUI',
			'    Set candidate = live.Item(1)',
			'    Debug.Print candidate.ToastSlot',
			'End Sub',
			'',
		].join('\n');

		expect(byCode(analyze(src), CODE)).toHaveLength(0);
	});

	it('stays silent when the name is also a Public member somewhere', () => {
		// Another class exposes the same name publicly, so a late-bound receiver
		// could legally dispatch there. The runtime type is unknowable; say nothing.
		const other = [
			'Option Explicit',
			'Public Property Get ToastSlot() As Long',
			'End Property',
			'',
		].join('\n');
		const src = [
			'Option Explicit',
			'Public Sub T()',
			'    Dim item As Variant',
			'    Debug.Print item.ToastSlot',
			'End Sub',
			'',
		].join('\n');

		const hits = byCode(analyze(src, [
			{ moduleName: 'OtherThing', moduleKind: 'class', source: other },
		]), CODE);
		expect(hits).toHaveLength(0);
	});

	it('stays silent for a member name that exists in the host object model', () => {
		// A Friend member named Value collides with the host surface, so a
		// late-bound receiver may well be a Range.
		const shadowing = [
			'Option Explicit',
			'Friend Property Get Value() As Long',
			'End Property',
			'',
		].join('\n');
		const src = [
			'Option Explicit',
			'Public Sub T()',
			'    Dim item As Variant',
			'    Debug.Print item.Value',
			'End Sub',
			'',
		].join('\n');

		const hits = byCode(analyzeProjectModule(src, [
			{ moduleName: 'Shadowing', moduleKind: 'class', source: shadowing },
		], 'Caller'), CODE);
		expect(hits).toHaveLength(0);
	});

	it('stays silent for an unknown member on a late-bound receiver', () => {
		// The VBE oracle records this as compile-valid: the runtime type could
		// support anything, so an unknown name is never evidence of a defect.
		const src = [
			'Option Explicit',
			'Public Sub T()',
			'    Dim item As Variant',
			'    Debug.Print item.NoSuchMemberAnywhere',
			'End Sub',
			'',
		].join('\n');

		expect(byCode(analyze(src), CODE)).toHaveLength(0);
	});

	it('stays silent on a strongly typed receiver of the owning class', () => {
		// Friend is legal within the project when the call is early bound.
		const src = [
			'Option Explicit',
			'Public Sub T()',
			'    Dim ui As ReDimUI',
			'    Set ui = New ReDimUI',
			'    Debug.Print ui.ToastSlot',
			'End Sub',
			'',
		].join('\n');

		expect(byCode(analyze(src), CODE)).toHaveLength(0);
	});

	it('stays silent with no project class context at all', () => {
		const src = [
			'Option Explicit',
			'Public Sub T()',
			'    Dim item As Variant',
			'    Debug.Print item.ToastSlot',
			'End Sub',
			'',
		].join('\n');

		expect(byCode(analyzeProjectModule(src, [], 'Caller'), CODE)).toHaveLength(0);
	});

	it('fires inside the owning class itself', () => {
		// Being inside the class does not put Friend members back on the
		// dispatch interface, so an Object-typed local fails here too. Found in
		// real code and confirmed by the VBE oracle
		// (late_bound_friend_member_same_class_runtime).
		const src = [
			'Option Explicit',
			'Public Function Probe(ByVal value As Variant) As Long',
			'    Dim candidate As Object',
			'    Set candidate = value',
			'    If TypeOf candidate Is ReDimUI Then',
			'        Probe = candidate.ToastSlot',
			'    End If',
			'End Function',
			'',
		].join('\n');

		const selfClass = UI_CLASS + src.replace('Option Explicit\n', '');
		const hits = byCode(analyzeProjectModule(selfClass, [
			{ moduleName: 'ReDimUI', moduleKind: 'class', source: selfClass },
		], 'ReDimUI'), CODE);

		expect(hits).toHaveLength(1);
		expect(spanText(selfClass, hits[0])).toBe('ToastSlot');
	});

	it('does not fire on a Collection member itself', () => {
		const src = [
			'Option Explicit',
			'Public Sub T(ByVal live As Collection)',
			'    Debug.Print live.Count',
			'End Sub',
			'',
		].join('\n');

		expect(byCode(analyze(src), CODE)).toHaveLength(0);
	});
});
