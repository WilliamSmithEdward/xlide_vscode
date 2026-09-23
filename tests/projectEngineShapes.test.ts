import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { ProjectEngine } from '../src/projectEngine';
import { listShapes } from '../src/vba/projectService';

// The shape parameters as they arrive from the agent tool and the editor:
// JSON, checked at the engine before anything reaches a file. The markup
// each look writes is covered per host in xlsxShapes, pptShapes and
// docShapes.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

let dir: string;
let deck: string;
const engine = new ProjectEngine({} as never);

beforeEach(() => {
	dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-engine-shapes-'));
	deck = path.join(dir, 'Deck.pptm');
	fs.copyFileSync(path.join(FIXTURES, 'PowerPointShapesFixture.pptm'), deck);
});

afterEach(() => {
	fs.rmSync(dir, { recursive: true, force: true });
});

const shape = (name: string) => listShapes(deck, 'Slide 1').surfaces[0].shapes.find((s) => s.name === name);

describe('shape parameters at the engine', () => {
	it('carries a whole look to the file, numbers given as text included', async () => {
		await engine.call('editShape', {
			path: deck, surface: 'Slide 1', action: 'update', name: 'ClickMe',
			fill: { type: 'solid', color: '#FF0000', transparency: 25 },
			line: { type: 'solid', color: '#008000', weight: '2.5', dash: 'dash' },
			font: { name: 'Arial', size: 14, bold: true, underline: false, color: '' },
			rotation: '30', hidden: true, zOrder: 'front',
		});

		expect(shape('ClickMe')).toMatchObject({
			rotation: 30, hidden: true, zOrder: 3,
			fill: { type: 'solid', color: '#FF0000', transparency: 25 },
			line: { type: 'solid', color: '#008000', weight: 2.5, dash: 'dash' },
			font: { name: 'Arial', size: 14, bold: true, underline: false, color: '#FFFFFF' },
		});
	});

	it('refuses a malformed look before it reaches the file', async () => {
		const before = fs.readFileSync(deck);
		const refusals: Array<[Record<string, unknown>, RegExp]> = [
			[{ fill: 'red' }, /'fill' parameter must be an object/],
			[{ fill: { type: 'gradient' } }, /fill's type must be none or solid, not 'gradient'/],
			[{ fill: { type: 'solid' } }, /A solid fill needs a color/],
			[{ line: { type: 'solid', dash: 'wavy' } }, /'wavy' is not a dash style/],
			[{ line: { type: 'solid', weight: 'thick' } }, /'weight' parameter must be a number of points, not 'thick'/],
			[{ font: { size: 'big' } }, /'size' parameter must be a number of points/],
			[{ font: { bold: 'yes' } }, /'bold' parameter must be true or false/],
			[{ zOrder: 'top' }, /'zOrder' parameter must be front, back, forward or backward, not 'top'/],
			[{ hidden: 'yes' }, /'hidden' parameter must be true or false/],
			[{ rotation: 'abc' }, /'rotation' parameter must be a number of degrees/],
		];
		for (const [look, message] of refusals) {
			await expect(engine.call('editShape', { path: deck, surface: 'Slide 1', action: 'update', name: 'ClickMe', ...look }), JSON.stringify(look))
				.rejects.toThrow(message);
		}
		expect(fs.readFileSync(deck).equals(before)).toBe(true);
	});

	it('lists the Subs a shape can run', async () => {
		const result = await engine.call<{ macros: Array<{ macro: string }> }>('shapeMacros', { path: deck });
		expect(result.macros.map((m) => m.macro)).toEqual(['SayHello', 'Unlinked']);
	});
});
