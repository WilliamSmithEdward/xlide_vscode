import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const referencePath = path.join(root, 'reference', 'excel', 'json', 'Workbook.json');
const outputPath = path.join(
	root,
	'src',
	'analyzer',
	'host',
	'excelWorkbookReferenceMembers.ts',
);
const jsonDir = path.dirname(referencePath);

const primitiveTypes = new Set([
	'Boolean',
	'Byte',
	'Currency',
	'Date',
	'Decimal',
	'Double',
	'Integer',
	'Long',
	'LongLong',
	'LongPtr',
	'Object',
	'Single',
	'String',
	'Variant',
	'void',
	'IUnknown',
]);

function readWorkbookDump() {
	return JSON.parse(fs.readFileSync(referencePath, 'utf8'));
}

function normalizeTypeName(typeName) {
	if (!typeName || typeof typeName !== 'string') {
		return undefined;
	}
	return typeName.replace(/\(.*\)$/g, '').trim();
}

function excelQualifiedReturn(typeName) {
	const clean = normalizeTypeName(typeName);
	if (!clean || primitiveTypes.has(clean)) {
		return undefined;
	}
	if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(clean)) {
		return undefined;
	}
	if (!fs.existsSync(path.join(jsonDir, `${clean}.json`))) {
		return undefined;
	}
	return `Excel.${clean}`;
}

function memberFrom(raw, kind) {
	const returns = kind === 'method' ? raw.returns : raw.type ?? raw.returns;
	const qualifiedReturn = excelQualifiedReturn(returns);
	return {
		name: raw.name,
		kind,
		...(qualifiedReturn ? { returns: qualifiedReturn } : {}),
	};
}

function collectMembers(workbook) {
	const byName = new Map();
	const add = (raw, kind) => {
		if (!raw?.name) {
			return;
		}
		const key = raw.name.toLowerCase();
		if (byName.has(key)) {
			return;
		}
		byName.set(key, memberFrom(raw, kind));
	};

	for (const item of workbook.properties ?? []) {
		add(item, 'property');
	}
	for (const item of workbook.methods ?? []) {
		add(item, 'method');
	}
	for (const item of workbook.events ?? []) {
		add(item, 'event');
	}
	return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name, 'en'));
}

function renderMember(member) {
	const parts = [
		`name: ${JSON.stringify(member.name)}`,
		`kind: ${JSON.stringify(member.kind)}`,
	];
	if (member.returns) {
		parts.push(`returns: ${JSON.stringify(member.returns)}`);
	}
	return `\t{ ${parts.join(', ')} },`;
}

function renderOutput(workbook, members) {
	const provenance = `${workbook.library}; ${workbook.guid}; reference/excel/json/Workbook.json`;
	return `// Generated from reference/excel/json/Workbook.json. Do not hand-edit member names here.
// Regenerate from the repo-local reference dump when the dump changes.

import type { HostMember } from './excelObjectModel';

export const EXCEL_WORKBOOK_REFERENCE_PROVENANCE = ${JSON.stringify(provenance)};

export const EXCEL_WORKBOOK_REFERENCE_MEMBERS: readonly HostMember[] = [
${members.map(renderMember).join('\n')}
];
`;
}

const workbook = readWorkbookDump();
const members = collectMembers(workbook);
fs.writeFileSync(outputPath, renderOutput(workbook, members), 'utf8');
console.log(`Generated ${members.length} Excel.Workbook members at ${outputPath}`);
