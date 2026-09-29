// Reads a YARA-X run and writes its matches as SARIF, for sarif-gate.mjs and
// the release report. Fails when the scan itself did not hold up: a file it
// could not scan, too few of the rules loaded, or the canary gone unmatched.
//   node .github/scripts/yara-x-sarif.mjs --matches <yr scan json>
//     --scan-log <its stderr> --loaded <yr scan --negate json of an empty file>
//     --compile-log <its stderr> --rules-name <text> --list <scan list>
//     --canary <file> --sarif <out> <rule file or directory>...
import fs from 'node:fs';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { CANARY_RULE, isMain, listedFiles, stripAnsi, toSarif } from './scan-sarif.mjs';

/**
 * The share of the rules that must load. YARA-X rejects some rules written
 * for YARA; far fewer loading means the rule package or the engine changed.
 */
export const MIN_RULES_LOADED = 0.9;

/** Each rule YARA-X reported, with the file it matched. */
export function readMatches(json) {
	return (json.matches ?? []).map((match) => ({
		rule: match.rule,
		file: match.file,
		description: typeof match.meta?.description === 'string' ? match.meta.description : undefined,
	}));
}

/** Lines of `yr` output that report a file it could not scan. */
export function scanErrors(text) {
	return stripAnsi(text).split(/\r?\n/).filter((line) => line.startsWith('error: '));
}

/** Rules YARA-X could not compile, each reported as `error[Ennn]: ...`. */
export function compileErrorCount(text) {
	return stripAnsi(text).split(/\r?\n/).filter((line) => /^error\[/.test(line)).length;
}

/** Rules declared in YARA source text. */
export function countRules(text) {
	return (text.match(/^[ \t]*(?:(?:private|global)[ \t]+)*rule[ \t]+[A-Za-z_][A-Za-z0-9_]*/gm) ?? []).length;
}

function ruleFiles(target) {
	if (fs.statSync(target).isFile()) {
		return [target];
	}
	return fs.readdirSync(target, { recursive: true })
		.map((name) => path.join(target, String(name)))
		.filter((file) => /\.(yar|yara)$/.test(file) && fs.statSync(file).isFile());
}

/**
 * A scanned file's path relative to the workspace, as the scan list named it:
 * yr prints the absolute path of each file it reads from a list.
 */
export function workspacePath(file, root) {
	const posix = file.replace(/\\/g, '/');
	const relative = posix.startsWith('/') ? path.posix.relative(root.replace(/\\/g, '/'), posix) : posix;
	return relative.replace(/^\.\//, '');
}

/** The findings to gate on, and every reason the scan cannot be trusted. */
export function judgeYaraX({ matches, scanLog, rulesLoaded, rulesInSource, canary, root }) {
	const problems = scanErrors(scanLog);
	if (rulesLoaded < rulesInSource * MIN_RULES_LOADED) {
		problems.push(`YARA-X loaded ${rulesLoaded} of ${rulesInSource} rules, fewer than ${MIN_RULES_LOADED * 100}%`);
	}
	const located = matches.map((match) => ({ ...match, file: workspacePath(match.file, root) }));
	const isCanary = (match) => match.file === workspacePath(canary, root);
	if (!located.some((match) => match.rule === CANARY_RULE && isCanary(match))) {
		problems.push(`YARA-X did not match the canary, ${canary}`);
	}
	const findings = located
		.filter((match) => !isCanary(match))
		.map((match) => ({
			ruleId: match.rule,
			uri: match.file,
			message: match.description ? `${match.rule}: ${match.description}` : match.rule,
		}));
	return { findings, problems };
}

if (isMain(import.meta)) {
	const { values, positionals } = parseArgs({
		allowPositionals: true,
		options: Object.fromEntries(['matches', 'scan-log', 'loaded', 'compile-log', 'rules-name', 'list', 'canary', 'sarif']
			.map((name) => [name, { type: 'string' }])),
	});
	const report = JSON.parse(fs.readFileSync(values.matches, 'utf8'));
	const matches = readMatches(report);
	const rulesLoaded = readMatches(JSON.parse(fs.readFileSync(values.loaded, 'utf8'))).length;
	const rulesInSource = positionals.flatMap(ruleFiles)
		.reduce((total, file) => total + countRules(fs.readFileSync(file, 'utf8')), 0);
	const { findings, problems } = judgeYaraX({
		matches,
		scanLog: fs.readFileSync(values['scan-log'], 'utf8'),
		rulesLoaded,
		rulesInSource,
		canary: values.canary,
		root: process.cwd(),
	});
	const properties = {
		engine: report.version,
		rules: values['rules-name'],
		rulesLoaded,
		rulesInSource,
		rulesNotCompiled: compileErrorCount(fs.readFileSync(values['compile-log'], 'utf8')),
		listedFiles: listedFiles(fs.readFileSync(values.list, 'utf8')).length,
	};
	const driver = { name: 'YARA-X', version: report.version, informationUri: 'https://virustotal.github.io/yara-x/' };
	fs.writeFileSync(values.sarif, `${JSON.stringify(toSarif(driver, findings, properties), null, 2)}\n`);
	console.log(`YARA-X ${report.version} with ${values['rules-name']}: ${rulesLoaded} of ${rulesInSource} rules loaded,`
		+ ` ${properties.listedFiles} files scanned, ${findings.length} match(es)`);
	for (const problem of problems) {
		console.error(`::error::${problem}`);
	}
	if (problems.length > 0) {
		process.exit(1);
	}
}
