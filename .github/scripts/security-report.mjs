// Writes the security report attached to a release: each gate's verdict and
// every finding the SARIF and npm audit output hold.
//   node .github/scripts/security-report.mjs <artifacts directory> <out.md>
// Reads CODEQL_RESULT, SEMGREP_RESULT, AUDIT_RESULT and TAG from the
// environment, and GITHUB_SHA and GITHUB_REPOSITORY as Actions sets them.
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

/** Results reviewed as not vulnerabilities (.github/codeql/reviewed.json), by rule and file. */
const reviewed = fs.existsSync('.github/codeql/reviewed.json')
	? JSON.parse(fs.readFileSync('.github/codeql/reviewed.json', 'utf8')).reviewed.map((entry) => ({
		rule: entry.rule,
		matches: new RegExp(`^${entry.file.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[^/]*')}$`),
	}))
	: [];

function sarifFindings(root) {
	const found = [];
	let tools = new Set();
	for (const file of files(root, '.sarif')) {
		const sarif = JSON.parse(fs.readFileSync(file, 'utf8'));
		for (const run of sarif.runs ?? []) {
			const driver = run.tool?.driver;
			if (driver) {
				tools.add(`${driver.name} ${driver.semanticVersion ?? driver.version ?? ''}`.trim());
			}
			for (const result of run.results ?? []) {
				const where = result.locations?.[0]?.physicalLocation;
				const uri = where?.artifactLocation?.uri ?? '?';
				const known = reviewed.some((entry) => entry.rule === result.ruleId && entry.matches.test(uri));
				found.push(`- \`${result.ruleId}\` at ${uri}:${where?.region?.startLine ?? '?'}${known ? ' (reviewed, see .github/codeql/reviewed.json)' : ''}`);
			}
		}
	}
	return { found, tools: [...tools] };
}

const verdict = (result) => (result === 'success' ? 'passed' : result ? `did not pass (${result})` : 'did not run');
const codeql = sarifFindings(path.join(directory, 'codeql-javascript-typescript'));
const codeqlActions = sarifFindings(path.join(directory, 'codeql-actions'));
const semgrep = sarifFindings(path.join(directory, 'semgrep'));
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
	['CodeQL, JavaScript and TypeScript', codeql.found],
	['CodeQL, GitHub Actions', codeqlActions.found],
	['Semgrep', semgrep.found],
];
for (const [title, found] of sections) {
	lines.push(`## ${title}`, '', ...(found.length ? found : ['No findings.']), '');
}
lines.push('## npm audit', '', '```', JSON.stringify(audit, null, 2), '```', '');
fs.writeFileSync(out, lines.join('\n'));
console.log(lines.join('\n'));
