// Reads a clamscan run and writes its detections as SARIF, for sarif-gate.mjs
// and the release report. Fails when the scan itself did not hold up: an
// error, counts that do not add up, or the canary gone undetected.
//   node .github/scripts/clamav-sarif.mjs --output <clamscan output>
//     --exit-code <n> --version <clamscan --version output>
//     --list <scan list> --canary <file> --sarif <out>
import fs from 'node:fs';
import { parseArgs } from 'node:util';
import { isMain, listedFiles, toSarif } from './scan-sarif.mjs';

const SUMMARY_FIELDS = ['Known viruses', 'Engine version', 'Scanned files', 'Infected files', 'Total errors'];

/** The detections, the problems and the summary clamscan printed. */
export function readClamscan(text) {
	const detections = [];
	const problems = [];
	const summary = {};
	for (const line of text.split(/\r?\n/)) {
		// A signature name has no spaces; a file name can.
		const found = /^(.*): (\S+) FOUND$/.exec(line);
		if (found) {
			detections.push({ file: found[1], signature: found[2] });
			continue;
		}
		const field = /^([A-Za-z ]+): (.*)$/.exec(line);
		if (field && SUMMARY_FIELDS.includes(field[1])) {
			summary[field[1]] = field[2].trim();
			continue;
		}
		// A file clamscan could not read, or could not finish.
		if (/ ERROR$/.test(line) || /^(ERROR|WARNING): /.test(line)) {
			problems.push(line);
		}
	}
	return { detections, problems, summary };
}

/** `ClamAV 1.5.4/27791/Sun Sep 28 07:35:12 2026`: the engine and the daily signatures. */
export function readVersion(text) {
	const [engine, database, databaseDate] = text.trim().replace(/^ClamAV /, '').split('/');
	return { engine, database, databaseDate };
}

const samePath = (a, b) => a.replace(/^\.\//, '') === b.replace(/^\.\//, '');

/**
 * The findings to gate on, and every reason the scan cannot be trusted.
 * clamscan exits 1 when it detects anything, and the canary is always there
 * to detect, so a sound run exits 1.
 */
export function judgeClamscan(scan, exitCode, canary) {
	const problems = [...scan.problems];
	if (exitCode !== 1) {
		problems.push(`clamscan exited with ${exitCode}; a sound run exits 1, having detected the canary`);
	}
	if (Number(scan.summary['Total errors'] ?? 0) > 0) {
		problems.push(`clamscan counted ${scan.summary['Total errors']} error(s)`);
	}
	if (!(Number(scan.summary['Scanned files']) > 0)) {
		problems.push('clamscan printed no count of scanned files');
	}
	if (Number(scan.summary['Infected files']) !== scan.detections.length) {
		problems.push(`clamscan counted ${scan.summary['Infected files']} infected file(s) but printed ${scan.detections.length}`);
	}
	if (!scan.detections.some((detection) => samePath(detection.file, canary))) {
		problems.push(`clamscan did not detect the canary, ${canary}`);
	}
	const findings = scan.detections
		.filter((detection) => !samePath(detection.file, canary))
		.map((detection) => ({
			ruleId: detection.signature,
			uri: detection.file.replace(/^\.\//, ''),
			message: `ClamAV reports ${detection.signature}`,
		}));
	return { findings, problems };
}

if (isMain(import.meta)) {
	const { values } = parseArgs({
		options: Object.fromEntries(['output', 'exit-code', 'version', 'list', 'canary', 'sarif']
			.map((name) => [name, { type: 'string' }])),
	});
	const scan = readClamscan(fs.readFileSync(values.output, 'utf8'));
	const version = readVersion(fs.readFileSync(values.version, 'utf8'));
	const { findings, problems } = judgeClamscan(scan, Number(values['exit-code']), values.canary);
	const properties = {
		engine: version.engine,
		database: version.database,
		databaseDate: version.databaseDate,
		knownSignatures: Number(scan.summary['Known viruses']),
		scannedFiles: Number(scan.summary['Scanned files']),
		listedFiles: listedFiles(fs.readFileSync(values.list, 'utf8')).length,
	};
	const driver = { name: 'ClamAV', version: version.engine, informationUri: 'https://www.clamav.net' };
	fs.writeFileSync(values.sarif, `${JSON.stringify(toSarif(driver, findings, properties), null, 2)}\n`);
	console.log(`ClamAV ${version.engine}, daily signatures ${version.database} (${version.databaseDate}):`
		+ ` ${properties.scannedFiles} of ${properties.listedFiles} listed files scanned, ${findings.length} detection(s)`);
	for (const problem of problems) {
		console.error(`::error::${problem}`);
	}
	if (problems.length > 0) {
		process.exit(1);
	}
}
