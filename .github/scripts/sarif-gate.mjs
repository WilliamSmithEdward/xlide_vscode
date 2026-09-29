// Fails when any SARIF file under a directory holds a result.
//   node .github/scripts/sarif-gate.mjs <label> <directory>
import fs from 'node:fs';
import path from 'node:path';

const [label, directory] = process.argv.slice(2);
const results = [];
for (const name of fs.existsSync(directory) ? fs.readdirSync(directory) : []) {
	if (!name.endsWith('.sarif')) {
		continue;
	}
	const sarif = JSON.parse(fs.readFileSync(path.join(directory, name), 'utf8'));
	for (const run of sarif.runs ?? []) {
		for (const result of run.results ?? []) {
			const where = result.locations?.[0]?.physicalLocation;
			results.push(`${result.ruleId}: ${where?.artifactLocation?.uri ?? '?'}:${where?.region?.startLine ?? '?'} ${result.message?.text ?? ''}`);
		}
	}
}
console.log(`${label}: ${results.length} result(s)`);
for (const line of results) {
	console.log(`  ${line}`);
}
if (results.length > 0) {
	process.exit(1);
}
