import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import {
	applyFormDesignerOp,
	applyFormMarkup,
	readFormMarkup,
	resetWorkbookCacheForTests,
} from '../src/vba/workbookService';

// The two orders a form carries that geometry does not show.
//
// DEPTH is the sibling order in the saved stream, LAST site on top - measured
// by letting MSForms do it: calling ZOrder(fmZOrderFront) on a control in the
// live designer and saving moved that control to the END of the site list,
// while its TabIndex stayed put (2026-08-31). TAB ORDER is the TabIndex run,
// renumbered as one gesture the way the VBE's dialog does.

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

const tempDirs: string[] = [];
afterEach(() => {
	resetWorkbookCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function workbook(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-order-'));
	tempDirs.push(dir);
	const wb = path.join(dir, 'Forms.xlsm');
	fs.copyFileSync(FIXTURE, wb);
	return wb;
}

/** Top-level control names in document order, which IS the depth order. */
function topLevelOrder(wb: string): string[] {
	const markup = readFormMarkup(wb, 'EntryForm').markup;
	resetWorkbookCacheForTests();
	return [...markup.matchAll(/^    <\w+ Name="([^"]+)"/gm)].map((m) => m[1]);
}

function tabIndexOf(wb: string, name: string): number | undefined {
	const markup = readFormMarkup(wb, 'EntryForm').markup;
	resetWorkbookCacheForTests();
	const line = markup.split(/\r\n/).find((l) => l.includes(`Name="${name}"`));
	const found = line ? /TabIndex="(-?\d+)"/.exec(line) : null;
	return found ? Number(found[1]) : undefined;
}

describe('z-order', () => {
	it('brings a control to the front by moving it last, and sends it back by moving it first', () => {
		const wb = workbook();
		const before = topLevelOrder(wb);
		expect(before.length).toBeGreaterThan(3);
		const first = before[0];

		applyFormDesignerOp(wb, 'EntryForm', { kind: 'zOrder', name: first, toFront: true });
		resetWorkbookCacheForTests();
		const front = topLevelOrder(wb);
		expect(front[front.length - 1]).toBe(first);
		// Everything else keeps its relative order.
		expect(front.filter((n) => n !== first)).toEqual(before.filter((n) => n !== first));

		applyFormDesignerOp(wb, 'EntryForm', { kind: 'zOrder', name: first, toFront: false });
		resetWorkbookCacheForTests();
		expect(topLevelOrder(wb)).toEqual(before);
	});

	it('leaves the tab order alone - depth and tab order are independent', () => {
		const wb = workbook();
		const name = topLevelOrder(wb)[0];
		const tab = tabIndexOf(wb, name);
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'zOrder', name, toFront: true });
		resetWorkbookCacheForTests();
		expect(tabIndexOf(wb, name)).toBe(tab);
	});

	it('is a no-op at the end it already occupies, and refuses a page', () => {
		const wb = workbook();
		const order = topLevelOrder(wb);
		const last = order[order.length - 1];
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'zOrder', name: last, toFront: true });
		resetWorkbookCacheForTests();
		expect(topLevelOrder(wb)).toEqual(order);
		expect(() => applyFormDesignerOp(wb, 'EntryForm', { kind: 'zOrder', name: 'Page1', toFront: true }))
			.toThrow(/Page/);
	});

	it('survives the save: the document carries the depth it set', () => {
		const real = workbook();
		const scratch = path.join(path.dirname(real), 'Scratch.xlsm');
		fs.copyFileSync(real, scratch);
		resetWorkbookCacheForTests();
		const name = topLevelOrder(scratch)[0];
		applyFormDesignerOp(scratch, 'EntryForm', { kind: 'zOrder', name, toFront: true });
		resetWorkbookCacheForTests();
		const doc = readFormMarkup(scratch, 'EntryForm').markup;
		resetWorkbookCacheForTests();
		applyFormMarkup(real, 'EntryForm', doc);
		resetWorkbookCacheForTests();
		expect(readFormMarkup(real, 'EntryForm').markup).toBe(doc);
		const after = topLevelOrder(real);
		expect(after[after.length - 1]).toBe(name);
	});
});

describe('tab order', () => {
	it('renumbers a whole surface from the order it is given', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		resetWorkbookCacheForTests();
		// Everything at the top level that carries a tab index.
		const named = [...markup.matchAll(/^    <\w+ Name="([^"]+)"[^>]*TabIndex="(\d+)"/gm)]
			.map((m) => ({ name: m[1], tab: Number(m[2]) }))
			.sort((a, b) => a.tab - b.tab);
		expect(named.length).toBeGreaterThan(3);

		const reversed = named.map((n) => n.name).reverse();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'tabOrder', container: '', names: reversed });
		resetWorkbookCacheForTests();
		reversed.forEach((name, index) => {
			expect(tabIndexOf(wb, name), `${name} at ${index}`).toBe(index);
		});
	});

	it('refuses a partial or unknown order rather than half-applying it', () => {
		const wb = workbook();
		const before = readFormMarkup(wb, 'EntryForm').markup;
		resetWorkbookCacheForTests();
		expect(() => applyFormDesignerOp(wb, 'EntryForm', { kind: 'tabOrder', container: '', names: ['NameBox'] }))
			.toThrow(/must list every control/);
		expect(() => applyFormDesignerOp(wb, 'EntryForm', { kind: 'tabOrder', container: '', names: ['NoSuchThing'] }))
			.toThrow(/not a control/);
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup).toBe(before);
	});

	it('orders a Frame-s own children, not the form-s', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		resetWorkbookCacheForTests();
		const frame = /<Frame Name="Options"[\s\S]*?<\/Frame>/.exec(markup)?.[0] ?? '';
		const kids = [...frame.matchAll(/<\w+ Name="([^"]+)"[^>]*TabIndex="(\d+)"/g)]
			.map((m) => ({ name: m[1], tab: Number(m[2]) }))
			.filter((k) => k.name !== 'Options')
			.sort((a, b) => a.tab - b.tab);
		expect(kids.length).toBeGreaterThan(1);
		const reversed = kids.map((k) => k.name).reverse();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'tabOrder', container: 'Options', names: reversed });
		resetWorkbookCacheForTests();
		reversed.forEach((name, index) => {
			expect(tabIndexOf(wb, name), `${name} at ${index}`).toBe(index);
		});
	});
});
