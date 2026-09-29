// Writes the security report attached to a release: each gate's verdict and
// every finding the SARIF and npm audit output hold.
//   node .github/scripts/security-report.mjs <artifacts directory> <out.md>
// Reads CODEQL_RESULT, SEMGREP_RESULT, AUDIT_RESULT, PACKAGE_RESULT,
// CLAMAV_RESULT, YARA_X_RESULT and TAG from the environment, and GITHUB_SHA
// and GITHUB_REPOSITORY as Actions sets them.
import fs from 'node:fs';
import path from 'node:path';

const [directory, out] = process.argv.slice(2);

function files(root, suffix) {
	if (!fs.existsSync(root)) {
		return [];
	}
	return fs.readdirSync(root, { recursive: true })
		.map((name) => path.join(root, String(name)))
		.filter((name) => name.endsWith(suffix));
}

/** Results reviewed and found harmless, by rule and file, from one reviewed list. */
function reviewedList(file) {
	const entries = fs.existsSync(file)
		? JSON.parse(fs.readFileSync(file, 'utf8')).reviewed.map((entry) => ({
			rule: entry.rule,
			matches: new RegExp(`^${entry.file.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[^/]*')}$`),
		}))
		: [];
	return { file, entries };
}
const codeReviewed = reviewedList('.github/codeql/reviewed.json');
const scanReviewed = reviewedList('.github/scans/reviewed.json');

function sarifFindings(root, reviewed) {
	const found = [];
	const tools = new Set();
	let properties = {};
	for (const file of files(root, '.sarif')) {
		const sarif = JSON.parse(fs.readFileSync(file, 'utf8'));
		for (const run of sarif.runs ?? []) {
			const driver = run.tool?.driver;
			if (driver) {
				tools.add(`${driver.name} ${driver.semanticVersion ?? driver.version ?? ''}`.trim());
			}
			properties = { ...properties, ...run.properties };
			for (const result of run.results ?? []) {
				const where = result.locations?.[0]?.physicalLocation;
				const uri = where?.artifactLocation?.uri ?? '?';
				const known = reviewed.entries.some((entry) => entry.rule === result.ruleId && entry.matches.test(uri));
				const line = where?.region?.startLine ? `:${where.region.startLine}` : '';
				found.push(`- \`${result.ruleId}\` at ${uri}${line}${known ? ` (reviewed, see ${reviewed.file})` : ''}`);
			}
		}
	}
	return { found, tools: [...tools], properties };
}

const verdict = (result) => (result === 'success' ? 'passed' : result ? `did not pass (${result})` : 'did not run');
const count = (value) => (typeof value === 'number' ? value.toLocaleString('en-US') : '?');
const codeql = sarifFindings(path.join(directory, 'codeql-javascript-typescript'), codeReviewed);
const codeqlActions = sarifFindings(path.join(directory, 'codeql-actions'), codeReviewed);
const semgrep = sarifFindings(path.join(directory, 'semgrep'), codeReviewed);
const clamav = sarifFindings(path.join(directory, 'clamav'), scanReviewed);
const yaraX = sarifFindings(path.join(directory, 'yara-x'), scanReviewed);
let audit = { total: '?' };
const auditFile = path.join(directory, 'npm-audit', 'npm-audit.json');
if (fs.existsSync(auditFile)) {
	audit = JSON.parse(fs.readFileSync(auditFile, 'utf8')).metadata?.vulnerabilities ?? audit;
}
const digestFile = path.join(directory, 'package-sha256', 'package-sha256.txt');
const [packageDigest, packagePath] = fs.existsSync(digestFile)
	? fs.readFileSync(digestFile, 'utf8').trim().split(/\s+/)
	: [];

const lines = [
	`# Security report for ${process.env.TAG ?? '(no tag)'}`,
	'',
	`Commit ${process.env.GITHUB_SHA ?? '?'} of ${process.env.GITHUB_REPOSITORY ?? '?'}.`,
	'Every gate fails on any finding; see SECURITY.md for what each covers.',
	'',
	'| Check | Verdict | Findings |',
	'| --- | --- | --- |',
	`| CodeQL (security-extended) | ${verdict(process.env.CODEQL_RESULT)} | ${codeql.found.length + codeqlActions.found.length} |`,
	`| Semgrep (p/default, p/typescript, p/github-actions) | ${verdict(process.env.SEMGREP_RESULT)} | ${semgrep.found.length} |`,
	`| npm audit (production and development) | ${verdict(process.env.AUDIT_RESULT)} | ${audit.total} |`,
	`| ClamAV | ${verdict(process.env.CLAMAV_RESULT)} | ${clamav.found.length} |`,
	`| YARA-X (${yaraX.properties.rules ?? 'YARA Forge core rules'}) | ${verdict(process.env.YARA_X_RESULT)} | ${yaraX.found.length} |`,
	'',
	`Tools: ${[...codeql.tools, ...codeqlActions.tools, ...semgrep.tools, ...clamav.tools, ...yaraX.tools].join(', ') || 'none recorded'}.`,
	'',
	'## What ClamAV and YARA-X scanned',
	'',
	packagePath
		? `\`${path.basename(packagePath)}\`, the file attached to this release (SHA-256 \`${packageDigest}\`),`
			+ ' with every file in the repository at this commit and every dependency npm installs from the'
			+ ` lockfile: ${count(clamav.properties.listedFiles ?? yaraX.properties.listedFiles)} files.`
		: `No package was scanned: the package job ${verdict(process.env.PACKAGE_RESULT)}.`,
	'',
	`- ClamAV ${clamav.properties.engine ?? '?'} with daily signatures ${clamav.properties.database ?? '?'}`
		+ ` (${clamav.properties.databaseDate ?? '?'}), ${count(clamav.properties.knownSignatures)} signatures in all.`,
	`- YARA-X ${yaraX.properties.engine ?? '?'} with ${yaraX.properties.rules ?? 'YARA Forge core rules'}:`
		+ ` ${count(yaraX.properties.rulesLoaded)} of ${count(yaraX.properties.rulesInSource)} rules loaded.`,
	'- Each scan also had to detect the EICAR test file written for the run, which shows its signatures'
		+ ' loaded and its output was read.',
	'',
];
const sections = [
	['CodeQL, JavaScript and TypeScript', codeql.found],
	['CodeQL, GitHub Actions', codeqlActions.found],
	['Semgrep', semgrep.found],
	['ClamAV', clamav.found],
	['YARA-X', yaraX.found],
];
for (const [title, found] of sections) {
	lines.push(`## ${title}`, '', ...(found.length ? found : ['No findings.']), '');
}
lines.push('## npm audit', '', '```', JSON.stringify(audit, null, 2), '```', '');
fs.writeFileSync(out, lines.join('\n'));
console.log(lines.join('\n'));
