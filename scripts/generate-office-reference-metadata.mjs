import fs from 'node:fs';
import path from 'node:path';
import {
	collectConstants,
	collectEnums,
	readReferenceDumps,
	renderConstant,
	renderEnum,
} from './reference-generator-utils.mjs';

const root = process.cwd();
const jsonDir = path.join(root, 'reference', 'office', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'officeReferenceConstants.ts');

const dumps = readReferenceDumps(jsonDir);
const constants = collectConstants(dumps);
const enums = collectEnums(dumps);

function renderOutput() {
	const constantEntries = constants.map(renderConstant).join('\n');
	const enumEntries = enums.map(renderEnum).join('\n');
	return `// Generated from reference/office/json. Do not hand-edit constants here.
// Regenerate from the repo-local reference dump with \`npm run generate:reference:office\`.

import type { HostConstant, HostEnum } from './excelObjectModel';

export const OFFICE_REFERENCE_ENUM_CONSTANTS: Record<string, HostConstant> = {
${constantEntries}
};

export const OFFICE_REFERENCE_ENUMS: Record<string, HostEnum> = {
${enumEntries}
};
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, renderOutput(), 'utf8');

console.log(
	`Generated ${constants.length} Office enum constant(s) in ${enums.length} enumeration(s) at ${outputPath}`,
);
