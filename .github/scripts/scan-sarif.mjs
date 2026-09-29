// What the ClamAV and YARA-X converters share: the SARIF they write, which
// sarif-gate.mjs and security-report.mjs read like any other, and the canary
// every scan must find.
//   node .github/scripts/scan-sarif.mjs canary <file>   writes the canary
import fs from 'node:fs';
import { pathToFileURL } from 'node:url';

/**
 * The EICAR test file, as hex so the repository never holds the file itself:
 * antivirus tools treat a file that starts with it as a detection. Each scan
 * writes it at run time and must report it, which proves the files were
 * scanned and the output was read.
 */
export const EICAR_HEX = '58354F2150254041505B345C505A58353428505E2937434329377D24'
	+ '45494341522D5354414E444152442D414E544956495255532D544553542D46494C4521'
	+ '24482B482A';

/** The canary rule's name in .github/scans/canary.yar. */
export const CANARY_RULE = 'XLIDE_SCAN_CANARY_EICAR';

export function writeCanary(file) {
	fs.writeFileSync(file, Buffer.from(EICAR_HEX, 'hex'));
}

/** Text with terminal colour codes removed. */
export function stripAnsi(text) {
	return text.replace(/\u001b\[[0-9;]*m/g, '');
}

/** The lines of a list file that name a file. */
export function listedFiles(text) {
	return text.split(/\r?\n/).filter((line) => line.length > 0);
}

/**
 * One SARIF run: `findings` are `{ ruleId, uri, message }`. The properties
 * record what was scanned and with which signatures, for the report.
 */
export function toSarif(driver, findings, properties) {
	const ruleIds = [...new Set(findings.map((finding) => finding.ruleId))].sort();
	return {
		$schema: 'https://json.schemastore.org/sarif-2.1.0.json',
		version: '2.1.0',
		runs: [{
			tool: {
				driver: {
					...driver,
					rules: ruleIds.map((id) => ({ id, shortDescription: { text: id } })),
				},
			},
			properties,
			results: findings.map((finding) => ({
				ruleId: finding.ruleId,
				level: 'error',
				message: { text: finding.message },
				locations: [{ physicalLocation: { artifactLocation: { uri: finding.uri } } }],
			})),
		}],
	};
}

/** True when the module is the script node was asked to run. */
export function isMain(meta) {
	return Boolean(process.argv[1]) && meta.url === pathToFileURL(process.argv[1]).href;
}

if (isMain(import.meta)) {
	const [command, file] = process.argv.slice(2);
	if (command !== 'canary' || !file) {
		console.error('usage: node .github/scripts/scan-sarif.mjs canary <file>');
		process.exit(2);
	}
	writeCanary(file);
}
