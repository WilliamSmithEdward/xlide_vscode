import fs from 'node:fs';
import path from 'node:path';
import {
	collectConstants,
	collectEnums,
	readReferenceDumps,
	renderConstant,
	renderEnum,
} from './reference-generator-utils.mjs';

// The DAO library (Microsoft Office 16.0 Access Database Engine Object
// Library, {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28} 12.0) is referenced by every
// new Access database, and most Access code names its constants:
// dbFailOnError, dbOpenDynaset, dbText (issue #103). The dump comes from
// pyVBAReference's scraper pointed at that library (reference/dao/json).
const root = process.cwd();
const jsonDir = path.join(root, 'reference', 'dao', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'daoReferenceConstants.ts');

const dumps = readReferenceDumps(jsonDir);
const constants = collectConstants(dumps);
const enums = collectEnums(dumps);

function renderOutput() {
	const constantEntries = constants.map(renderConstant).join('\n');
	const enumEntries = enums.map(renderEnum).join('\n');
	return `// Generated from reference/dao/json. Do not hand-edit constants here.
// Regenerate from the repo-local reference dump with \`npm run generate:reference:dao\`.
//
// The DAO library (Microsoft Office 16.0 Access Database Engine Object Library)
// is referenced by every new Access database, so its constants are everyday
// names in Access code: dbFailOnError, dbOpenDynaset, dbText (issue #103).

import type { HostConstant, HostEnum } from './excelObjectModel';

export const DAO_REFERENCE_ENUM_CONSTANTS: Record<string, HostConstant> = {
${constantEntries}
};

export const DAO_REFERENCE_ENUMS: Record<string, HostEnum> = {
${enumEntries}
};
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, renderOutput(), 'utf8');

console.log(
	`Generated ${constants.length} DAO enum constant(s) in ${enums.length} enumeration(s) at ${outputPath}`,
);
