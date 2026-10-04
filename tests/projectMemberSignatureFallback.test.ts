import { expect, it } from 'vitest';
import { resolveMemberCompletionNamed, resolveMemberCompletions, type MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const source = 'Sub Go()\nDim actor As Class1\nactor.\nEnd Sub\n';
const offset = source.indexOf('actor.') + 6;
function member(name: string, signature?: string): VbaProjectClassMember {
	return { name, moduleName: 'Class1', kind: 'property', returns: 'Object', signature };
}
function context(members: VbaProjectClassMember[], kind: VbaProjectClassMembers['kind'] = 'class'): MemberCompletionContext {
	return { projectClassMembers: [{ name: 'Class1', moduleName: 'Class1', kind, exhaustive: true, members }], memberSurfaceCache: new Map() };
}
function expected(name: string, signature?: string) {
	return { name, kind: 'property', returns: 'Object', signature, declaredType: undefined, access: undefined,
		writable: undefined, writeType: undefined, owner: 'Class1', surfaceExhaustive: true, documentation: undefined,
		doc: undefined, definitions: undefined, defaultMember: undefined, letAccessor: undefined, sub: undefined,
		setAccessor: undefined, attributes: undefined };
}
for (const position of ['first', 'last']) {
	it.each([10, 100, 1000])('bounds absent-signature named lookup at %i members, ' + position, count => {
		let reads = 0;
		const ctx = context(Array.from({ length: count }, (_, i) => ({ ...member('M' + i), get name() { reads++; return 'M' + i; } })));
		const name = 'M' + (position === 'first' ? 0 : count - 1);
		for (let i = 0; i < count; i++) expect(resolveMemberCompletionNamed(source, offset, name, ctx)).toEqual(expected(name));
		expect(reads).toBeLessThanOrEqual(count * 5);
	});
}
it.each([10, 100, 1000])('bounds completion menu absent-signature lookup at %i members', count => {
	let reads = 0;
	const ctx = context(Array.from({ length: count }, (_, i) => ({ ...member('M' + i), get name() { reads++; return 'M' + i; } })));
	expect(resolveMemberCompletions(source, offset, ctx)).toEqual(Array.from({ length: count }, (_, i) => expected('M' + i)));
	expect(reads).toBeLessThanOrEqual(count * 5);
});
it.each(['class', 'document', 'userType', 'enum', 'standardModule'] as const)('preserves %s signature selection', kind => {
	const ctx = context([member('First'), member('Second', 'Second() As Object')], kind);
	expect(resolveMemberCompletions(source, offset, ctx)).toEqual(kind === 'enum' || kind === 'standardModule' ? [] : [expected('First'), expected('Second', 'Second() As Object')]);
});
it.each([undefined, 'FIRST() As Object', ''])('preserves first case-insensitive signature %j', signature => {
	for (const cached of [false, true]) {
		const ctx = context([member('First', signature), member('FIRST', 'Later() As Object')]);
		if (!cached) delete ctx.memberSurfaceCache;
		expect(resolveMemberCompletionNamed(source, offset, 'fIrSt', ctx)).toEqual(expected('First', signature));
	}
});
it('refreshes absent and supplied signatures between contexts with retained metadata', () => {
	const members = [member('First')];
	expect(resolveMemberCompletionNamed(source, offset, 'First', context(members))).toEqual(expected('First'));
	members[0].signature = 'First() As Object';
	expect(resolveMemberCompletionNamed(source, offset, 'First', context(members))).toEqual(expected('First', 'First() As Object'));
	delete members[0].signature;
	expect(resolveMemberCompletionNamed(source, offset, 'First', context(members))).toEqual(expected('First'));
});
