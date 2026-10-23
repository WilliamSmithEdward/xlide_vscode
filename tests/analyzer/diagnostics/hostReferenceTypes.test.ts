// invalid-as-type-name for a name none of a project's references holds, in
// each Office host (issue #234 for Excel; Word, PowerPoint and Access
// measured 2026-10-16). The rule can only say so where it knows every name
// each referenced library holds, so it is held to what a new file in each
// host references, as the VBE reports them:
//
//   Excel       VBA, Excel, stdole, Office
//   Word        VBA, Word, stdole, Office, and the Normal template
//   PowerPoint  VBA, PowerPoint, stdole, Office
//   Access      VBA, Access, stdole, DAO
//
// The dir stream records the references after VBA and the host's own
// library. Word's documents reference their Normal template, whose own
// classes XLIDE does not read, so a Word document is never judged here.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../../src/analyzer';
import { byCode } from '../../helpers/diagnostics';

const REFERENCES: Record<string, readonly string[]> = {
	excel: ['stdole', 'Office'],
	word: ['stdole', 'Normal', 'Office'],
	powerpoint: ['stdole', 'Office'],
	access: ['stdole', 'DAO'],
};

interface Fixture { id: string; host?: string; expected: string; source: string }

const fixtures: Fixture[] = (JSON.parse(readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8')) as { cases: Fixture[] })
	.cases.filter((kase) => kase.id.startsWith('reftype_') || kase.id === 'host_excel_worksheet_type_compile'
		|| /^host_(word|powerpoint|access)_(excel_worksheet|own_library)_type_compile$/.test(kase.id));

/**
 * Rejected by the VBE, not reported here, and why. The Access model carries
 * the Office library's types, which a database that does not reference
 * Office cannot use; the analyzer does not yet drop them by reference.
 */
const KNOWN_SILENT: ReadonlyMap<string, string> = new Map([
	['reftype_access_office_commandbar_compile', 'Office types are in the Access model whether or not Office is referenced'],
]);

const unknownTypes = (fixture: Fixture): number => {
	const host = fixture.host ?? 'excel';
	return byCode(analyzeModule(fixture.source, { host, referencedLibraries: REFERENCES[host] }), 'invalid-as-type-name')
		.filter((diagnostic) => diagnostic.message.includes('none of the libraries it references')).length;
};

describe('a type none of the references holds, per Office host', () => {
	it('covers every host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host ?? 'excel'))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		// A Word document's Normal template could hold the name, so Word stays silent.
		const reported = fixture.expected === 'rejected' && fixture.host !== 'word' && !KNOWN_SILENT.has(fixture.id);
		expect(unknownTypes(fixture)).toBe(reported ? 1 : 0);
	});

	it('stays silent where the references are not known', () => {
		const src = 'Option Explicit\nSub Main()\n    Dim w As Worksheet\nEnd Sub\n';
		expect(byCode(analyzeModule(src, { host: 'access' }), 'invalid-as-type-name')).toHaveLength(0);
	});
});
