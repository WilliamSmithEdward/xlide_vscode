// What a database's designs say about their members (issue #206). The design's
// TypeInfo stream is the list the compiler checks against; a design without
// one has members nobody can list.

import { readFileSync } from 'fs';
import * as path from 'path';
import { afterEach, describe, expect, it, vi } from 'vitest';

let dropTypeInfo = false;
vi.mock('../src/vba/access/accessVbaWriter', async (importOriginal) => {
	const actual = await importOriginal<typeof import('../src/vba/access/accessVbaWriter')>();
	return {
		...actual,
		readAccessDesignNames: (data: Buffer) => actual.readAccessDesignNames(data)
			.map((design) => (dropTypeInfo ? { ...design, typeInfo: undefined } : design)),
	};
});

import { openMacroContainer } from '../src/vba/macroContainer';

const DATABASE = path.join(__dirname, 'fixtures', 'binaries', 'AccessBoundFormFixture.accdb');

afterEach(() => {
	dropTypeInfo = false;
});

function ordersMembers(): { name: string; type: string }[] | undefined {
	const design = openMacroContainer(readFileSync(DATABASE)).designs?.().find((entry) => entry.name === 'Orders');
	return design?.members(1252);
}

describe('the members of an Access design', () => {
	it('are the TypeInfo stream s, record-source fields included', () => {
		expect(ordersMembers()?.map((member) => member.name)).toEqual(['Detail', 'Qty', 'Order_Date', 'ID', 'Amount', 'Unit Price']);
	});

	it('are not known when the design has no TypeInfo stream, rather than its controls alone', () => {
		dropTypeInfo = true;
		expect(ordersMembers()).toBeUndefined();
	});
});
