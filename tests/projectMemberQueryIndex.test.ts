import { expect, it } from 'vitest';
import { projectClassMemberAt, resolveReceiverTypeAt, type MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
const source = 'Sub Go()\nDim actor As Class1\nactor.\nEnd Sub\n', offset = source.indexOf('actor.') + 6;
function member(name: string, returns = 'Object'): VbaProjectClassMember { return { name, moduleName: 'Class1', kind: 'property', returns, writable: true }; }
function surface(members: VbaProjectClassMember[], kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers { return { name: 'Class1', moduleName: 'Class1', kind, members, exhaustive: true }; }
function context(type: VbaProjectClassMembers, text = source): MemberCompletionContext { return { projectClassMembers: [type], parsedModule: parseModule(text), memberSurfaceCache: new Map() }; }
for (const position of ['first', 'last', 'distinct']) {
	it.each([10, 100, 1000])('bounds raw project query at %i members, ' + position, count => {
		let reads = 0;
		const members = Array.from({ length: count }, (_, i) => ({ ...member('M' + i), get name() { reads++; return 'M' + i; } })), ctx = context(surface(members));
		for (let i = 0; i < count; i++) { const at = position === 'first' ? 0 : position === 'last' ? count - 1 : i; expect(projectClassMemberAt(source, offset, 'M' + at, ctx)).toBe(members[at]); }
		expect(reads).toBeLessThanOrEqual(count);
	});
}
it.each([10, 100, 1000])('bounds full diagnostic raw member lookup at %i writes', count => {
	let reads = 0;
	const members = Array.from({ length: count }, (_, i) => ({ ...member('M' + i), get name() { reads++; return 'M' + i; } }));
	const text = ['Option Explicit', 'Sub Go()', 'Dim actor As New Class1', ...Array.from({ length: count }, (_, i) => 'actor.M' + i + ' = 1'), 'End Sub', ''].join('\n');
	let from = 0;
	const expected = Array.from({ length: count }, (_, i) => { const label = 'actor.M' + i, start = text.indexOf(label, from) + 6; from = start + label.length - 6; return ['set-required', "Object assignment to '" + label + "' requires Set because it expects Object.", { start, end: start + label.length - 6 }]; });
	const errors: unknown[] = [];
	expect(analyzeModule(text, { projectClassMembers: [surface(members)], onInternalError: e => errors.push(e) }).map(d => [d.code, d.message, d.span])).toEqual(expected);
	expect(errors).toEqual([]); expect(reads).toBeLessThanOrEqual(count * 10 + 100);
});
it.each(['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const)('preserves raw %s eligibility', kind => {
	const m = member('M'), ctx = context(surface([m], kind));
	expect(projectClassMemberAt(source, offset, 'M', ctx)).toBe(kind === 'class' ? m : undefined);
});
it.each([undefined, 'M() As Object', ''])('keeps the first duplicate and its full metadata, signature=%j', signature => {
	const first = { ...member('M'), signature, sub: true }, second = { ...member('m', 'Long'), signature: 'Later() As Long' }, ctx = context(surface([first, second]));
	expect(projectClassMemberAt(source, offset, 'm', ctx)).toBe(first);
});
it('resumes the index after a first hit and caches misses without rescanning', () => {
	let reads = 0;
	const members = Array.from({ length: 1000 }, (_, i) => ({ ...member('M' + i), get name() { reads++; return 'M' + i; } })), ctx = context(surface(members));
	expect(projectClassMemberAt(source, offset, 'M0', ctx)).toBe(members[0]); expect(reads).toBe(1);
	for (let i = 0; i < 100; i++) expect(projectClassMemberAt(source, offset, 'Missing', ctx)).toBeUndefined();
	expect(reads).toBe(1000); expect(projectClassMemberAt(source, offset, 'M999', ctx)).toBe(members[999]); expect(reads).toBe(1000);
});
it('keeps uncached queries cheap and fresh for retained mutable metadata', () => {
	let reads = 0;
	const first = { ...member('First'), get name() { reads++; return 'First'; } }, members = [first, ...Array.from({ length: 1000 }, (_, i) => member('M' + i))], ctx = context(surface(members));
	delete ctx.memberSurfaceCache;
	expect(projectClassMemberAt(source, offset, 'First', ctx)).toBe(first); expect(reads).toBe(1);
	expect(projectClassMemberAt(source, offset, 'New', ctx)).toBeUndefined();
	const added = member('New'); members.push(added); expect(projectClassMemberAt(source, offset, 'New', ctx)).toBe(added);
});
it('refreshes retained member metadata with a fresh pass cache', () => {
	const members = [member('First')], type = surface(members);
	expect(projectClassMemberAt(source, offset, 'New', context(type))).toBeUndefined();
	const added = member('New'); members.push(added);
	expect(projectClassMemberAt(source, offset, 'New', context(type))).toBe(added);
	members.splice(1, 1); expect(projectClassMemberAt(source, offset, 'New', context(type))).toBeUndefined();
});
it('does not promote implicit controls into raw project members', () => {
	const text = source.replace('actor.', 'Me.'), ctx = { ...context(surface([]), text), meProjectType: 'Class1', implicitMembers: [{ name: 'Implicit', type: 'MSForms.TextBox' }] };
	expect(projectClassMemberAt(text, text.indexOf('Me.') + 3, 'Implicit', ctx)).toBeUndefined();
});
it('preserves cached project receiver return chains with argument-sensitive signatures', () => {
	const text = 'Sub Go()\nDim actor As Class1\nactor.M(1).\nEnd Sub\n', ctx = context(surface([member('M', 'Class2')]), text);
	ctx.projectClassMembers = [...ctx.projectClassMembers!, { name: 'Class2', moduleName: 'Class2', kind: 'class', members: [] }];
	expect(resolveReceiverTypeAt(text, text.indexOf('actor.M(1).') + 11, ctx)).toBe('project:class2');
});
