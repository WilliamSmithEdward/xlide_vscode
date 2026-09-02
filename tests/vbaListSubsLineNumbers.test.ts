import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as svc from '../src/vba/projectService';

// Procedure line numbers used to be computed per match with
// `body.slice(0, m.index).split('\n').length`, which re-walked the module from
// the start for every procedure found. That is quadratic: on the 26,000-line
// ROneCOne class, expanding the module in the explorer cost about 370 ms, of
// which the search itself was 1.5 ms. One forward pass counts each character
// once. Widely separated procedures are what a broken running counter gets
// wrong, so that is what this pins.
const TEMPLATE = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');

function tempWorkbook(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-lines-'));
	const target = path.join(dir, 'Book.xlsm');
	fs.copyFileSync(TEMPLATE, target);
	return target;
}

describe('procedure line numbers', () => {
	it('stay correct when procedures are thousands of lines apart', () => {
		const target = tempWorkbook();
		const filler = Array.from({ length: 500 }, () => '    Debug.Print 1').join('\r\n');
		const source = [
			'Public Sub First()',
			filler,
			'End Sub',
			'',
			'Public Sub Second()',
			filler,
			'End Sub',
			'',
			'Public Function Third() As Long',
			'End Function',
			'',
		].join('\r\n');
		svc.writeModule(target, 'Spread', source, 'standard');

		const subs = svc.listSubs(target, 'Spread');
		expect(subs.map((s) => s.name)).toEqual(['First', 'Second', 'Third']);

		// Every reported line must actually contain that procedure's header in
		// the body the editor shows - the check the old formula also satisfied.
		const body = svc.readModule(target, 'Spread', false).source.split('\r\n');
		for (const sub of subs) {
			expect(body[sub.line - 1], sub.name).toContain(sub.name);
		}

		// And the exact numbers: 500 filler lines + End Sub + blank between each.
		expect(subs.map((s) => s.line)).toEqual([1, 504, 1007]);
	});

	it('counts the first procedure as line 1', () => {
		const target = tempWorkbook();
		svc.writeModule(target, 'AtTop', 'Public Sub Only()\r\nEnd Sub\r\n', 'standard');
		expect(svc.listSubs(target, 'AtTop')).toEqual([{ name: 'Only', kind: 'Sub', line: 1 }]);
	});

	it('agrees with a direct scan on a module with many procedures', () => {
		const target = tempWorkbook();
		const parts: string[] = [];
		for (let i = 0; i < 200; i += 1) {
			parts.push(`Public Sub P${i}()`, '    Debug.Print 1', 'End Sub', '');
		}
		svc.writeModule(target, 'Many', parts.join('\r\n'), 'standard');

		const subs = svc.listSubs(target, 'Many');
		expect(subs).toHaveLength(200);
		const body = svc.readModule(target, 'Many', false).source.split('\r\n');
		for (const sub of subs) {
			expect(body[sub.line - 1], sub.name).toContain(`Sub ${sub.name}(`);
		}
	});
});
