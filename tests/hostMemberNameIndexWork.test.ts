import {describe, expect, it, vi} from 'vitest';
import {isHostMemberName, isHostMemberNameAnywhere} from '../src/analyzer/host/hostModel';
import type {HostObjectModel, HostType} from '../src/analyzer/host/excelObjectModel';

function modelWith(count: number): HostObjectModel {
	const types: Record<string, HostType> = {};
	for (let i = 0; i < count; i++) types['Test.Type' + i] = Object.freeze({displayName: 'Type' + i, members: Object.freeze([
		Object.freeze({name: `member_${i}_property`, kind: 'property' as const}),
		Object.freeze({name: `member_${i}_method`, kind: 'method' as const}),
		Object.freeze({name: `member_${i}_event`, kind: 'event' as const}),
	])});
	return Object.freeze({source: 'test', types: Object.freeze(types), aliases: Object.freeze({}), globals: Object.freeze({})});
}

describe('host broad member-name index reuse', () => {
	it.each([10, 1000])('does not rebuild %i type member names after the shared index is ready', count => {
		const model = modelWith(count);
		expect(isHostMemberNameAnywhere('member_0_event', model)).toBe(true);
		let adds = 0;
		const original = Set.prototype.add;
		const spy = vi.spyOn(Set.prototype, 'add').mockImplementation(function(this: Set<unknown>, value: unknown) {
			if (typeof value === 'string' && value.startsWith('member_')) adds++;
			return original.call(this, value);
		});
		let result;
		try {result = isHostMemberName(`MEMBER_${count - 1}_METHOD`, model);} finally {spy.mockRestore();}
		expect(result).toBe(true);
		expect(adds).toBe(0);
	});

	it('includes events, properties and methods, but not globals or unknown names', () => {
		const model = Object.freeze({...modelWith(1), globals: Object.freeze({GlobalOnly: 'Test.Type0'})});
		for (const suffix of ['event', 'property', 'method']) {
			expect(isHostMemberName('MEMBER_0_' + suffix.toUpperCase(), model)).toBe(true);
		}
		expect(isHostMemberName('GlobalOnly', model)).toBe(false);
		expect(isHostMemberName('missing', model)).toBe(false);
		expect(isHostMemberName('', model)).toBe(false);
	});

	it('keeps separate model ownership and either lookup initialization order', () => {
		const first = modelWith(1), second = modelWith(2);
		expect(isHostMemberName('member_0_method', first)).toBe(true);
		expect(isHostMemberNameAnywhere('member_1_method', second)).toBe(true);
		expect(isHostMemberName('member_1_method', first)).toBe(false);
		expect(isHostMemberName('member_1_method', second)).toBe(true);
		expect(isHostMemberNameAnywhere('member_0_event', first)).toBe(true);
	});
});
