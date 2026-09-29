// Fails when any SARIF file under a directory holds a result that is not in
// the reviewed list (rule and file, `*` matching within one path segment).
//   node .github/scripts/sarif-gate.mjs <label> <directory> [reviewed.json]
import fs from 'node:fs';
import path from 'node:path';

const [label, directory, reviewedFile] = process.argv.slice(2);

/** The reviewed entries, each with a matcher for its file. */
export function loadReviewed(file) {
	if (!file || !fs.existsSync(file)) {
		return [];
	}
	return JSON.parse(fs.readFileSync(file, 'utf8')).reviewed.map((entry) => ({
		...entry,
		matches: new RegExp(`^${entry.file.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[^/]*')}$`),
	}));
}

export function isReviewed(reviewed, ruleId, uri) {
	return reviewed.some((entry) => entry.rule === ruleId && entry.matches.test(uri));
}

const reviewed = loadReviewed(reviewedFile);
const open = [];
let accepted = 0;
for (const name of fs.existsSync(directory) ? fs.readdirSync(directory) : []) {
	if (!name.endsWith('.sarif')) {
		continue;
	}
	const sarif = JSON.parse(fs.readFileSync(path.join(directory, name), 'utf8'));
	for (const run of sarif.runs ?? []) {
		for (const result of run.results ?? []) {
			const where = result.locations?.[0]?.physicalLocation;
			const uri = where?.artifactLocation?.uri ?? '?';
			if (isReviewed(reviewed, result.ruleId, uri)) {
				accepted += 1;
				continue;
			}
			open.push(`${result.ruleId}: ${uri}:${where?.region?.startLine ?? '?'} ${result.message?.text ?? ''}`);
		}
	}
}
console.log(`${label}: ${open.length} unreviewed result(s), ${accepted} reviewed`);
for (const line of open) {
	console.log(`  ${line}`);
}
if (open.length > 0) {
	process.exit(1);
}
