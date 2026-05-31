import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const jsonDir = path.join(root, 'reference', 'excel', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'excelReferenceMembers.ts');
const coveragePath = path.join(root, 'docs', 'excel_reference_coverage.md');

// Keep runtime promotion explicit. The generator scans the full Excel corpus for
// coverage, but only these types become checked-in host metadata.
const promotedTypes = ['Workbook'];

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

function readJson(filePath) {
	return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function referenceFiles() {
	return fs
		.readdirSync(jsonDir)
		.filter((name) => name.endsWith('.json') && name !== '_index.json')
		.sort((a, b) => a.localeCompare(b, 'en'));
}

const dumps = new Map(
	referenceFiles().map((fileName) => {
		const dump = readJson(path.join(jsonDir, fileName));
		return [dump.name, { dump, fileName }];
	}),
);

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
	if (!dumps.has(clean)) {
		return undefined;
	}
	return `Excel.${clean}`;
}

function memberReturn(raw, kind) {
	return kind === 'method' ? raw.returns : raw.type ?? raw.returns;
}

function memberFrom(raw, kind) {
	const qualifiedReturn = excelQualifiedReturn(memberReturn(raw, kind));
	return {
		name: raw.name,
		kind,
		...(qualifiedReturn ? { returns: qualifiedReturn } : {}),
	};
}

function collectMembers(typeDump) {
	const byName = new Map();
	const duplicateNames = new Set();
	const add = (raw, kind) => {
		if (!raw?.name) {
			return;
		}
		const key = raw.name.toLowerCase();
		if (byName.has(key)) {
			duplicateNames.add(raw.name);
			return;
		}
		byName.set(key, memberFrom(raw, kind));
	};

	for (const item of typeDump.properties ?? []) {
		add(item, 'property');
	}
	for (const item of typeDump.methods ?? []) {
		add(item, 'method');
	}
	for (const item of typeDump.events ?? []) {
		add(item, 'event');
	}
	return {
		members: [...byName.values()].sort((a, b) => a.name.localeCompare(b.name, 'en')),
		duplicateNames: [...duplicateNames].sort((a, b) => a.localeCompare(b, 'en')),
	};
}

function renderMember(member) {
	const parts = [
		`name: ${JSON.stringify(member.name)}`,
		`kind: ${JSON.stringify(member.kind)}`,
	];
	if (member.returns) {
		parts.push(`returns: ${JSON.stringify(member.returns)}`);
	}
	return `\t\t{ ${parts.join(', ')} },`;
}

function provenanceFor(typeName) {
	const entry = dumps.get(typeName);
	if (!entry) {
		throw new Error(`Missing Excel reference dump for promoted type ${typeName}`);
	}
	const { dump, fileName } = entry;
	return `${dump.library}; ${dump.guid}; reference/excel/json/${fileName}`;
}

function renderPromotedMemberSet(typeName) {
	const entry = dumps.get(typeName);
	if (!entry) {
		throw new Error(`Missing Excel reference dump for promoted type ${typeName}`);
	}
	const { members } = collectMembers(entry.dump);
	return `\t${JSON.stringify(typeName)}: [\n${members.map(renderMember).join('\n')}\n\t],`;
}

function typeCoverage(entry) {
	const { dump, fileName } = entry;
	const properties = dump.properties?.length ?? 0;
	const methods = dump.methods?.length ?? 0;
	const events = dump.events?.length ?? 0;
	const constants = dump.constants?.length ?? 0;
	const totalMembers = properties + methods + events;
	const { members, duplicateNames } = collectMembers(dump);
	const memberRows = [
		...(dump.properties ?? []).map((item) => ({ item, kind: 'property' })),
		...(dump.methods ?? []).map((item) => ({ item, kind: 'method' })),
		...(dump.events ?? []).map((item) => ({ item, kind: 'event' })),
	];
	const membersWithReturnType = memberRows.filter(({ item, kind }) =>
		memberReturn(item, kind),
	).length;
	const membersWithQualifiedReturnType = members.filter((member) => member.returns).length;
	const membersWithSignature = memberRows.filter(({ item }) => item.signature).length;
	const parameters = memberRows.reduce(
		(sum, { item }) => sum + (item.parameters?.length ?? 0),
		0,
	);
	const typedParameters = memberRows.reduce(
		(sum, { item }) =>
			sum + (item.parameters ?? []).filter((param) => Boolean(param.type)).length,
		0,
	);
	return {
		name: dump.name,
		kind: dump.kind ?? 'Unknown',
		fileName,
		properties,
		methods,
		events,
		constants,
		totalMembers,
		uniqueMembers: members.length,
		duplicateNames: duplicateNames.length,
		membersWithReturnType,
		membersWithQualifiedReturnType,
		membersWithSignature,
		parameters,
		typedParameters,
	};
}

const coverage = [...dumps.values()]
	.map(typeCoverage)
	.sort((a, b) => a.name.localeCompare(b.name, 'en'));

function countWhere(predicate) {
	return coverage.filter(predicate).length;
}

function sumOf(field, rows = coverage) {
	return rows.reduce((sum, row) => sum + row[field], 0);
}

function markdownTable(headers, rows) {
	return [
		`| ${headers.join(' | ')} |`,
		`| ${headers.map(() => '---').join(' | ')} |`,
		...rows.map((row) => `| ${row.join(' | ')} |`),
	].join('\n');
}

function renderCoverageMarkdown() {
	const objectRows = coverage.filter((row) => row.totalMembers > 0);
	const enumRows = coverage.filter((row) => row.constants > 0);
	const promotedRows = promotedTypes.map((typeName) => {
		const row = coverage.find((item) => item.name === typeName);
		if (!row) {
			throw new Error(`Missing coverage row for promoted type ${typeName}`);
		}
		return row;
	});
	const largestRows = [...objectRows]
		.sort(
			(a, b) =>
				b.uniqueMembers - a.uniqueMembers ||
				a.name.localeCompare(b.name, 'en'),
		)
		.slice(0, 30);

	return `# Excel Reference Coverage

Generated by \`npm run generate:reference:excel\` from \`reference/excel/json\`.

## Summary

${markdownTable(
	['Metric', 'Count'],
	[
		['JSON files scanned', coverage.length],
		['Object-like type dumps', countWhere((row) => row.totalMembers > 0)],
		['Enumeration dumps', enumRows.length],
		['Promoted runtime types', promotedTypes.length],
		['Raw properties', sumOf('properties')],
		['Raw methods', sumOf('methods')],
		['Raw events', sumOf('events')],
		['Raw enum constants', sumOf('constants')],
		['Raw member rows', sumOf('totalMembers')],
		['Unique object member names', sumOf('uniqueMembers')],
		['Member rows with signatures', sumOf('membersWithSignature')],
		['Member rows with return/type data', sumOf('membersWithReturnType')],
		['Parameters with type data', sumOf('typedParameters')],
	],
)}

## Promoted Runtime Types

${markdownTable(
	[
		'Type',
		'Members',
		'Properties',
		'Methods',
		'Events',
		'Return/type rows',
		'Qualified returns',
		'Signatures',
		'Duplicate names',
	],
	promotedRows.map((row) => [
		row.name,
		row.uniqueMembers,
		row.properties,
		row.methods,
		row.events,
		row.membersWithReturnType,
		row.membersWithQualifiedReturnType,
		row.membersWithSignature,
		row.duplicateNames,
	]),
)}

## Largest Object Surfaces

${markdownTable(
	['Type', 'Kind', 'Members', 'Properties', 'Methods', 'Events', 'Signatures'],
	largestRows.map((row) => [
		row.name,
		row.kind,
		row.uniqueMembers,
		row.properties,
		row.methods,
		row.events,
		row.membersWithSignature,
	]),
)}

## Notes

- Runtime extension code does not read \`reference/\`; promoted metadata is checked in under \`src/\`.
- Completion may use partial metadata, but hard \`member-not-found\` diagnostics require a promoted exhaustive surface.
- Promotion remains type-by-type so each host surface can get representative tests and oracle controls before red diagnostics rely on absence.
`;
}

function renderOutput() {
	const provenanceEntries = promotedTypes
		.map((typeName) => `\t${JSON.stringify(typeName)}: ${JSON.stringify(provenanceFor(typeName))},`)
		.join('\n');
	const memberSetEntries = promotedTypes.map(renderPromotedMemberSet).join('\n');
	const workbookProvenance = provenanceFor('Workbook');

	return `// Generated from reference/excel/json. Do not hand-edit member names here.
// Regenerate from the repo-local reference dump with \`npm run generate:reference:excel\`.

import type { HostMember } from './excelObjectModel';

export const EXCEL_REFERENCE_PROMOTED_TYPES = ${JSON.stringify(promotedTypes)} as const;

export const EXCEL_REFERENCE_PROVENANCE: Record<string, string> = {
${provenanceEntries}
};

export const EXCEL_REFERENCE_MEMBER_SETS: Record<string, readonly HostMember[]> = {
${memberSetEntries}
};

export const EXCEL_WORKBOOK_REFERENCE_PROVENANCE = ${JSON.stringify(workbookProvenance)};

export const EXCEL_WORKBOOK_REFERENCE_MEMBERS = EXCEL_REFERENCE_MEMBER_SETS.Workbook;
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, renderOutput(), 'utf8');
fs.mkdirSync(path.dirname(coveragePath), { recursive: true });
fs.writeFileSync(coveragePath, renderCoverageMarkdown(), 'utf8');

console.log(
	`Generated ${promotedTypes.length} promoted Excel reference type(s) at ${outputPath}`,
);
console.log(`Wrote Excel reference coverage report at ${coveragePath}`);
