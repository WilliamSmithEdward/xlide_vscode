// The Office workflow (.github/workflows/office.yml) runs the VBE oracle on a
// GitHub-hosted Windows runner for the cases in runner_sample.json. A case
// renamed or removed from the corpus would only show there, after Office has
// been installed, so the list is checked here against the fixtures.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';

interface Case {
	id: string;
	host?: string;
	expected: string;
	mode: string;
	provenance?: string;
	crashesHostBuilds?: string[];
	references?: { guid: string }[];
}

const ORACLE = join(process.cwd(), 'syntax_corpus/oracle');
const corpus = JSON.parse(readFileSync(join(ORACLE, 'vbe_oracle_cases.json'), 'utf8')) as { cases: Case[] };
const sample = JSON.parse(readFileSync(join(ORACLE, 'runner_sample.json'), 'utf8')) as { cases: string[] };
const byId = new Map(corpus.cases.map((kase) => [kase.id, kase]));
const OUTLOOK_LIBRARY = '{00062FFF-0000-0000-C000-000000000046}';

describe('the runner sample', () => {
	it('names each case once, and only verified corpus cases that crash no build', () => {
		expect(new Set(sample.cases).size).toBe(sample.cases.length);
		for (const id of sample.cases) {
			const kase = byId.get(id);
			expect(kase, id).toBeDefined();
			expect(kase?.provenance, id).toBe('vbe-oracle-verified');
			expect(kase?.crashesHostBuilds, id).toBeUndefined();
		}
	});

	it.each(Object.keys(OFFICE_HOST_APPS))('has an accepted and a rejected compile and run in %s', (host) => {
		const found = sample.cases.map((id) => byId.get(id)).filter((kase) => (kase?.host ?? 'excel') === host)
			.map((kase) => `${kase?.mode} ${kase?.expected}`);
		for (const kind of ['compile accepted', 'compile rejected', 'run accepted', 'run rejected']) {
			expect(found).toContain(kind);
		}
	});

	it('has an accepted and a rejected case referencing the Outlook library, which the workflow installs Outlook for', () => {
		const outlook = sample.cases.map((id) => byId.get(id))
			.filter((kase) => kase?.references?.some((reference) => reference.guid === OUTLOOK_LIBRARY));
		expect(outlook.map((kase) => kase?.expected).sort()).toEqual(['accepted', 'rejected']);
	});

	it('is the list the workflow runs', () => {
		expect(readFileSync(join(process.cwd(), '.github/workflows/office.yml'), 'utf8'))
			.toMatch(/^\s*\$sample = \(Get-Content syntax_corpus\/oracle\/runner_sample\.json -Raw \| ConvertFrom-Json\)\.cases$/m);
	});
});
