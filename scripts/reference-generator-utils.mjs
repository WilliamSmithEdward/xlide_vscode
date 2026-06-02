import fs from 'node:fs';
import path from 'node:path';

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

export function constantFrom(raw, enumName) {
	if (!raw?.name) {
		return undefined;
	}
	return {
		name: raw.name,
		type: enumName,
		...(raw.value !== undefined ? { value: raw.value } : {}),
		source: 'external',
	};
}

export function collectConstants(dumps) {
	const byName = new Map();
	for (const { dump } of dumps.values()) {
		for (const raw of dump.constants ?? []) {
			const constant = constantFrom(raw, dump.name);
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

export function renderConstant(constant) {
	return `\t${JSON.stringify(constant.name)}: ${JSON.stringify(constant)},`;
}
