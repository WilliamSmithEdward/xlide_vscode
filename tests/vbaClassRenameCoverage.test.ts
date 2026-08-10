import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { projectClassReferenceLocations } from '../src/vbaNavigation';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';

// Issue #9 rule 4: renaming an interface has to follow it into `Implements`,
// into the `Interface_Member` prefix (which is a contract, not a coincidence),
// and into `As` / `New` type positions.
interface Mod { moduleName: string; source: string; type?: string }

const IFACE = 'Option Explicit\r\n\r\nPublic Sub Draw()\r\nEnd Sub\r\n';
const IMPL = [
	'Option Explicit',
	'',
	'Implements IShape',
	'',
	'Private Sub IShape_Draw()',
	'    Debug.Print "drawing"',
	'End Sub',
	'',
].join('\r\n');
const USER = [
	'Option Explicit',
	'',
	'Public Sub Use()',
	'    Dim s As IShape',
	'    Set s = New RoundShape',
	'    s.Draw',
	'    Dim IShapeLookalike As Long',   // must be left alone
	'    Debug.Print "IShape"',          // string: left alone
	"    ' IShape in a comment",         // comment: left alone
	'End Sub',
	'',
].join('\r\n');

function locationsFor(): { module: string; text: string }[] {
	const mods: Mod[] = [
		{ moduleName: 'IShape', source: IFACE, type: 'class' },
		{ moduleName: 'RoundShape', source: IMPL, type: 'class' },
		{ moduleName: 'Uses', source: USER },
	];
	const project = buildVbaProjectIndex(
		mods.map((m) => ({ moduleName: m.moduleName, source: m.source, type: m.type })),
	);
	const byModule = new Map(mods.map((m) => [m.moduleName.toLowerCase(), m as never]));
	const definition = project.projectTypeNames?.('Uses')?.find?.(
		(t: { name: string }) => t.name === 'IShape',
	);
	const locations = projectClassReferenceLocations(
		'C:/w/Book.xlsm', byModule as never, project, 'IShape',
		(definition ?? { name: 'IShape', moduleName: 'IShape', kind: 'class' }) as never,
	);
	return locations.map((loc) => {
		const uri = String((loc as { uri: { path?: string; toString(): string } }).uri);
		const module = /\/([^/]+)\.(?:cls|bas|frm)/.exec(uri)?.[1] ?? uri;
		const line = (loc as { range: { start: { line: number } } }).range.start.line;
		const source = mods.find((m) => m.moduleName === module)?.source ?? '';
		return { module, text: (source.split('\r\n')[line] ?? '').trim() };
	});
}

describe('renaming a class follows it everywhere it is named', () => {
	it('covers Implements, As-position and the Interface_Member prefix', () => {
		const found = locationsFor().map((hit) => `${hit.module}: ${hit.text}`);
		expect(found).toContain('RoundShape: Implements IShape');
		expect(found).toContain('Uses: Dim s As IShape');
		// The prefix is the contract; leaving it behind stops the class
		// implementing anything, which is a compile error the developer did
		// not ask for.
		expect(found).toContain('RoundShape: Private Sub IShape_Draw()');
	});

	it('leaves a lookalike name, a string and a comment alone', () => {
		const found = locationsFor().map((hit) => hit.text);
		expect(found.some((t) => t.includes('IShapeLookalike'))).toBe(false);
		expect(found.some((t) => t.startsWith('Debug.Print'))).toBe(false);
		expect(found.some((t) => t.startsWith("'"))).toBe(false);
	});
});
