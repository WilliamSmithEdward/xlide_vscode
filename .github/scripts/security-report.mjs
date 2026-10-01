// Writes the security report attached to a release: each gate's verdict and
// every finding the SARIF and npm audit output hold.
//   node .github/scripts/security-report.mjs <artifacts directory> <out.md>
// Reads CODEQL_RESULT, SEMGREP_RESULT, AUDIT_RESULT and TAG from the
// environment, and GITHUB_SHA and GITHUB_REPOSITORY as Actions sets them.
import fs from 'node:fs';
import path from 'node:path';
import { describeEntry, judge, loadReviewed, sarifResults } from './sarif-gate.mjs';

const [directory, out] = process.argv.slice(2);

function files(root, suffix) {
	if (!fs.existsSync(root)) {
		return [];
	}
	return fs.readdirSync(root, { recursive: true })
		.map((name) => path.join(root, String(name)))
		.filter((name) => name.endsWith(suffix));
}

/** Results reviewed as not vulnerabilities, matched as the gate matches them. */
const REVIEWED = '.github/codeql/reviewed.json';

function sarifFindings(root, scan) {
	let tools = new Set();
	const sarifFiles = files(root, '.sarif');
	for (const file of sarifFiles) {
		const sarif = JSON.parse(fs.readFileSync(file, 'utf8'));
		for (const run of sarif.runs ?? []) {
			const driver = run.tool?.driver;
			if (driver) {
				tools.add(`${driver.name} ${driver.semanticVersion ?? driver.version ?? ''}`.trim());
			}
		}
	}
	const results = sarifResults(sarifFiles);
	const { accepted, stale } = judge(results, loadReviewed(REVIEWED, scan));
	const found = results.map((result) =>
		`- \`${result.rule}\` at ${result.uri}:${result.line ?? '?'}${accepted.includes(result) ? ` (reviewed, see ${REVIEWED})` : ''}`);
	return {
		found,
		stale: stale.map((entry) => `- Reviewed in ${REVIEWED}, but no longer found: ${describeEntry(entry)}`),
		tools: [...tools],
	};
}

const verdict = (result) => (result === 'success' ? 'passed' : result ? `did not pass (${result})` : 'did not run');
const codeql = sarifFindings(path.join(directory, 'codeql-javascript-typescript'), 'CodeQL javascript-typescript');
const codeqlActions = sarifFindings(path.join(directory, 'codeql-actions'), 'CodeQL actions');
const semgrep = sarifFindings(path.join(directory, 'semgrep'), 'Semgrep');
let audit = { total: '?' };
const auditFile = path.join(directory, 'npm-audit', 'npm-audit.json');
if (fs.existsSync(auditFile)) {
	audit = JSON.parse(fs.readFileSync(auditFile, 'utf8')).metadata?.vulnerabilities ?? audit;
}

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
	'',
	`Tools: ${[...codeql.tools, ...codeqlActions.tools, ...semgrep.tools].join(', ') || 'none recorded'}.`,
	'',
];
const sections = [
	['CodeQL, JavaScript and TypeScript', codeql],
	['CodeQL, GitHub Actions', codeqlActions],
	['Semgrep', semgrep],
];
for (const [title, { found, stale }] of sections) {
	lines.push(`## ${title}`, '', ...(found.length ? found : ['No findings.']), ...stale, '');
}
lines.push('## npm audit', '', '```', JSON.stringify(audit, null, 2), '```', '');
fs.writeFileSync(out, lines.join('\n'));
console.log(lines.join('\n'));
