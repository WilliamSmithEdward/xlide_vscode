import { describe, expect, it } from 'vitest';
import { getHostMembers } from '../src/analyzer/host/hostModel';
import { resolveMemberCompletions } from '../src/analyzer';

// Issue #56: the model carries the type library's plumbing, and until now
// nothing marked it. A member is now flagged when the library says hidden or
// restricted, or when its name is one VBA cannot write at all - and a flagged
// member is still KNOWN (resolved, hovered, coloured) while never being
// OFFERED, because accepting the proposal would not compile.
describe('hidden host members', () => {
	it('flags the dispatch plumbing the VBE will not let you type', () => {
		const members = getHostMembers('Excel.Workbook');
		const codeName = members.find((m) => m.name === '_CodeName');
		expect(codeName, 'the model must still CARRY _CodeName').toBeDefined();
		expect(codeName?.hidden).toBe(true);
	});

	it('leaves the everyday members alone', () => {
		// Range.Rows / Columns / EntireRow are marked FNONBROWSABLE in the
		// library because they are duplicated on Application and Global.
		// Treating that flag as "hidden" would delete them from completion.
		const range = getHostMembers('Excel.Range');
		for (const name of ['Rows', 'Columns', 'EntireRow', 'EntireColumn', 'Value']) {
			const member = range.find((m) => m.name === name);
			expect(member, `Range.${name} must exist`).toBeDefined();
			expect(member?.hidden, `Range.${name} must not be hidden`).toBeFalsy();
		}
	});

	it('does not OFFER a hidden member, while still resolving it', () => {
		const source = 'Sub T()\r\n    ThisWorkbook.\r\nEnd Sub\r\n';
		const offered = resolveMemberCompletions(source, source.indexOf('ThisWorkbook.') + 'ThisWorkbook.'.length);
		expect(offered.length).toBeGreaterThan(20);
		expect(offered.some((c) => c.name === 'Worksheets')).toBe(true);
		expect(offered.some((c) => c.name.startsWith('_'))).toBe(false);
		expect(offered.some((c) => c.hidden)).toBe(false);
	});

	it('offers no member whose name VBA could not write', () => {
		// Whatever the receiver, a proposal must be something the VBE compiles.
		for (const receiver of ['ThisWorkbook.', 'Application.', 'ActiveSheet.']) {
			const source = `Sub T()\r\n    ${receiver}\r\nEnd Sub\r\n`;
			const offered = resolveMemberCompletions(source, source.indexOf(receiver) + receiver.length);
			const unwritable = offered.filter((c) => !/^[A-Za-z]/.test(c.name));
			expect(unwritable.map((c) => c.name), `${receiver} offered an unwritable name`).toEqual([]);
		}
	});
});
