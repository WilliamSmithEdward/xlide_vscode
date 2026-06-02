import fs from 'node:fs';
import path from 'node:path';
import {
	collectConstants,
	readReferenceDumps,
	renderConstant,
} from './reference-generator-utils.mjs';

const root = process.cwd();
const jsonDir = path.join(root, 'reference', 'office', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'officeReferenceConstants.ts');

const dumps = readReferenceDumps(jsonDir);
const constants = collectConstants(dumps);

function renderOutput() {
	const constantEntries = constants.map(renderConstant).join('\n');
	return `// Generated from reference/office/json. Do not hand-edit constants here.
// Regenerate from the repo-local reference dump with \`npm run generate:reference:office\`.

import type { HostConstant } from './excelObjectModel';

export const OFFICE_REFERENCE_ENUM_CONSTANTS: Record<string, HostConstant> = {
${constantEntries}
};
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, renderOutput(), 'utf8');

console.log(`Generated ${constants.length} Office enum constant(s) at ${outputPath}`);
