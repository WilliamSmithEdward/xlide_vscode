import fs from 'node:fs';
import path from 'node:path';
import { localizeHostName } from './reference-curation.mjs';

export function readJson(filePath) {
	return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

export function referenceFiles(jsonDir) {
	return fs
		.readdirSync(jsonDir)
		.filter((name) => name.endsWith('.json') && name !== '_index.json')
		.sort((a, b) => a.localeCompare(b, 'en'));
}

export function readReferenceDumps(jsonDir) {
	return new Map(
		referenceFiles(jsonDir).map((fileName) => {
			const dump = readJson(path.join(jsonDir, fileName));
			return [dump.name, { dump, fileName }];
		}),
	);
}

export function cleanText(value) {
	if (typeof value !== 'string') {
		return undefined;
	}
	const text = value.replace(/\s+/g, ' ').trim();
	return text.length > 0 ? text : undefined;
}

export function constantFrom(raw, enumName, hostApp = undefined, index = undefined) {
	if (!raw?.name) {
		return undefined;
	}
	// The reference describes 9,555 of the 11,847 enum members across these
	// libraries ("xlCategory: Axis displays categories"), and every generator
	// used to drop that prose on the floor.
	const summary = localizeHostName(cleanText(raw.description), hostApp, index);
	return {
		name: raw.name,
		type: enumName,
		...(raw.value !== undefined ? { value: raw.value } : {}),
		source: 'external',
		...(summary ? { doc: { summary, params: [], source: 'external' } } : {}),
	};
}

export function collectConstants(dumps, hostApp = undefined, index = undefined) {
	const byName = new Map();
	for (const { dump } of dumps.values()) {
		for (const raw of dump.constants ?? []) {
			const constant = constantFrom(raw, dump.name, hostApp, index);
			if (!constant) {
				continue;
			}
			const key = constant.name.toLowerCase();
			if (!byName.has(key)) {
				byName.set(key, constant);
			}
		}
	}
	return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name, 'en'));
}

/**
 * The enumerations a corpus declares, name and description. VBA accepts an enum
 * name as a declared type and as a qualifier, so the model needs the name even
 * though the members travel separately as constants.
 */
export function collectEnums(dumps, hostApp = undefined, index = undefined) {
	const byName = new Map();
	for (const { dump } of dumps.values()) {
		if (dump.kind !== 'Enumeration' || !dump.name || byName.has(dump.name)) {
			continue;
		}
		const summary = localizeHostName(cleanText(dump.description), hostApp, index);
		byName.set(dump.name, {
			displayName: dump.name,
			...(summary ? { doc: { summary, params: [], source: 'external' } } : {}),
		});
	}
	return [...byName.values()].sort((a, b) => a.displayName.localeCompare(b.displayName, 'en'));
}

export function renderEnum(entry) {
	return `	${JSON.stringify(entry.displayName)}: ${JSON.stringify(entry)},`;
}

export function renderConstant(constant) {
	return `\t${JSON.stringify(constant.name)}: ${JSON.stringify(constant)},`;
}
