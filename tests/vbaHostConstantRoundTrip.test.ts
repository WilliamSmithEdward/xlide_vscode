// Tripwire: the generated host models faithfully carry their reference dumps.
//
// Each host model's constants come from two sources: the host type library's
// Enumeration dumps (reference/<host>/json, via
// generate-host-object-model.mjs) and the shared Office library table
// (OFFICE_REFERENCE_ENUM_CONSTANTS, itself generated from
// reference/office/json). This suite walks every dump and asserts the model
// resolves every constant with the dumped value and enum type, and carries
// nothing the two sources do not - so a regeneration that drops, mangles, or
// invents constants announces itself.
//
// The reference corpus is deliberately NOT committed (.gitignore: a local
// transcription source, never bundled), so the dump-walking tests run only
// where the corpus exists - exactly the machines where regeneration can
// happen - and skip everywhere else (CI, fresh clones). The Office-table
// check needs no corpus and always runs.
import { describe, expect, it } from 'vitest';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import { getExcelObjectModel, type HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import { OFFICE_REFERENCE_ENUM_CONSTANTS } from '../src/analyzer/host/officeReferenceConstants';
import { MSFORMS_REFERENCE_ENUM_CONSTANTS } from '../src/analyzer/host/msformsReferenceMembers';
import { getHostConstants, resolveHostConstant } from '../src/analyzer/host/hostModel';

interface EnumDump {
	kind: string;
	name: string;
	constants?: Array<{ name: string; value?: unknown }>;
}

function dumpDirectory(host: string): string {
	return join(__dirname, '..', 'reference', host, 'json');
}

/** Lowercased name -> the values and enum types the host's dumps carry for it. */
function dumpedConstants(host: string): Map<string, { values: Set<string>; types: Set<string> }> {
	const dir = dumpDirectory(host);
	const byLower = new Map<string, { values: Set<string>; types: Set<string> }>();
	for (const file of readdirSync(dir)) {
		if (!file.endsWith('.json')) {
			continue;
		}
		let dump: EnumDump;
		try {
			dump = JSON.parse(readFileSync(join(dir, file), 'utf8')) as EnumDump;
		} catch {
			continue;
		}
		if (dump.kind !== 'Enumeration') {
			continue;
		}
		for (const constant of dump.constants ?? []) {
			if (constant.name.startsWith('_')) {
				continue;
			}
			const key = constant.name.toLowerCase();
			const entry = byLower.get(key) ?? { values: new Set<string>(), types: new Set<string>() };
			entry.values.add(
				typeof constant.value === 'number' ? String(constant.value) : String(constant.value ?? ''),
			);
			entry.types.add(dump.name);
			byLower.set(key, entry);
		}
	}
	return byLower;
}

const HOSTS: ReadonlyArray<[string, () => HostObjectModel]> = [
	['word', getWordObjectModel],
	['powerpoint', getPowerPointObjectModel],
	['access', getAccessObjectModel],
];

describe.each(HOSTS)('%s constants round-trip from the reference dumps', (host, getModel) => {
	const corpus = existsSync(dumpDirectory(host));

	it.runIf(corpus)('resolves every dumped enum constant with the dumped value and enum type', () => {
		const dumped = dumpedConstants(host);
		const model = getModel();
		expect(dumped.size).toBeGreaterThan(1000);
		const missing: string[] = [];
		const wrong: string[] = [];
		for (const [lower, entry] of dumped) {
			const resolved = resolveHostConstant(lower, model);
			if (!resolved) {
				missing.push(lower);
				continue;
			}
			if (!entry.values.has(String(resolved.value))) {
				wrong.push(`${resolved.name}: model=${String(resolved.value)} dump=${[...entry.values].join('|')}`);
			} else if (resolved.type && !entry.types.has(resolved.type)) {
				wrong.push(`${resolved.name}: model type=${resolved.type} dump=${[...entry.types].join('|')}`);
			}
		}
		expect(missing.slice(0, 10), `missing (${missing.length})`).toEqual([]);
		expect(wrong.slice(0, 10), `wrong (${wrong.length})`).toEqual([]);
	});

	it.runIf(corpus)('resolves every shared Office constant, host library winning shared names', () => {
		const dumped = dumpedConstants(host);
		const model = getModel();
		const wrong: string[] = [];
		for (const office of Object.values(OFFICE_REFERENCE_ENUM_CONSTANTS)) {
			const resolved = resolveHostConstant(office.name, model);
			if (!resolved) {
				wrong.push(`${office.name}: missing`);
				continue;
			}
			// A name the host library also exports answers with the host's dump
			// value; a purely shared name answers with the Office table's.
			const hostEntry = dumped.get(office.name.toLowerCase());
			const expected = hostEntry ? hostEntry.values : new Set([String(office.value)]);
			if (!expected.has(String(resolved.value))) {
				wrong.push(`${office.name}: model=${String(resolved.value)} expected=${[...expected].join('|')}`);
			}
		}
		expect(wrong.slice(0, 10), `wrong (${wrong.length})`).toEqual([]);
	});

	it.runIf(corpus)('carries nothing beyond the host dumps, Office and MSForms tables', () => {
		const dumped = dumpedConstants(host);
		const model = getModel();
		// The forms library joined the models in issue #41: fm* names are legal
		// wherever a project carries a UserForm, which is every host.
		const msforms = new Set(
			Object.values(MSFORMS_REFERENCE_ENUM_CONSTANTS).map((c) => c.name.toLowerCase()),
		);
		const invented = getHostConstants(model)
			.map((constant) => constant.name.toLowerCase())
			.filter((lower) => !dumped.has(lower) && !officeLowerNames().has(lower) && !msforms.has(lower));
		expect(invented.slice(0, 10), `invented (${invented.length})`).toEqual([]);
	});
});

let OFFICE_LOWER: Set<string> | undefined;
function officeLowerNames(): Set<string> {
	OFFICE_LOWER ??= new Set(
		Object.values(OFFICE_REFERENCE_ENUM_CONSTANTS).map((constant) => constant.name.toLowerCase()),
	);
	return OFFICE_LOWER;
}

it('the Excel model resolves every shared Office constant too', () => {
	const model = getExcelObjectModel();
	const missing = Object.values(OFFICE_REFERENCE_ENUM_CONSTANTS)
		.filter((office) => !resolveHostConstant(office.name, model))
		.map((office) => office.name);
	expect(missing.slice(0, 10), `missing (${missing.length})`).toEqual([]);
});
