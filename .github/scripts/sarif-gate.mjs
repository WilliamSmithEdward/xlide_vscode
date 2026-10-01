// Fails when any SARIF file under a directory holds a result that is not in
// the reviewed list, or when an entry of the list for this scan matches no
// result: an entry left behind would accept the next result to land there.
//   node .github/scripts/sarif-gate.mjs <scan> <directory> [reviewed.json]
// <scan> is one of SCANS, and an entry applies to the scan it names. A result
// that points at a line is matched by rule, exact path and the text of that
// line in the checkout, trimmed, so it follows code that moves and comes back
// for review when the line changes. A detection of a whole file is matched by
// rule and `file`, where `*` matches within one path segment.
import fs from 'node:fs';
import path from 'node:path';
import { isMain } from './scan-sarif.mjs';

/** The gates that read a reviewed list, as the workflows name them. */
export const SCANS = ['CodeQL javascript-typescript', 'CodeQL actions', 'ClamAV', 'YARA-X'];

function nonEmpty(value) {
	return typeof value === 'string' && value.trim() !== '';
}

/**
 * The entries of a reviewed list that apply to `scan`. A malformed entry, for
 * any scan, refuses the whole list, so a typo cannot accept a result.
 */
export function loadReviewed(file, scan) {
	if (!file || !fs.existsSync(file)) {
		return [];
	}
	const entries = JSON.parse(fs.readFileSync(file, 'utf8')).reviewed;
	if (!Array.isArray(entries)) {
		throw new Error(`${file}: "reviewed" is not a list`);
	}
	return entries.map((entry, index) => {
		const where = `${file}, entry ${index + 1}`;
		if (!SCANS.includes(entry.scan)) {
			throw new Error(`${where}: scan must be one of ${SCANS.join(', ')}`);
		}
		if (!nonEmpty(entry.rule) || !nonEmpty(entry.reason)) {
			throw new Error(`${where}: needs a rule and a reason`);
		}
		const byLine = nonEmpty(entry.path) && nonEmpty(entry.line) && entry.line === entry.line.trim();
		const byFile = nonEmpty(entry.file);
		if (byLine === byFile || (byFile && (entry.path !== undefined || entry.line !== undefined))) {
			throw new Error(`${where}: needs either path and the trimmed line, or file`);
		}
		return byFile
			? { ...entry, matches: new RegExp(`^${entry.file.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[^/]*')}$`) }
			: entry;
	}).filter((entry) => entry.scan === scan);
}

/** Line `line` of the file at `uri` under `root`, trimmed, or undefined when it cannot be read. */
export function lineText(root, uri, line) {
	if (!Number.isInteger(line) || line < 1) {
		return undefined;
	}
	const file = path.resolve(root, uri);
	const relative = path.relative(root, file);
	if (relative === '' || relative.startsWith('..') || path.isAbsolute(relative) || !fs.existsSync(file)) {
		return undefined;
	}
	return fs.readFileSync(file, 'utf8').split(/\r?\n/)[line - 1]?.trim();
}

/** Every result in the SARIF files, with the text of the line it flags read from the checkout at `root`. */
export function sarifResults(files, root = process.cwd()) {
	const results = [];
	for (const file of files) {
		for (const run of JSON.parse(fs.readFileSync(file, 'utf8')).runs ?? []) {
			for (const result of run.results ?? []) {
				const where = result.locations?.[0]?.physicalLocation;
				const uri = where?.artifactLocation?.uri ?? '?';
				const line = where?.region?.startLine;
				results.push({
					rule: result.ruleId,
					uri,
					line,
					text: line === undefined ? undefined : lineText(root, uri, line),
					message: result.message?.text ?? '',
				});
			}
		}
	}
	return results;
}

export function isReviewed(entry, result) {
	if (entry.rule !== result.rule) {
		return false;
	}
	if (result.line === undefined) {
		return entry.matches?.test(result.uri) ?? false;
	}
	return entry.path === result.uri && result.text !== undefined && entry.line === result.text;
}

/** The results no entry accepts, those accepted, and the entries no result matched. */
export function judge(results, reviewed) {
	const used = new Set();
	const open = [];
	const accepted = [];
	for (const result of results) {
		const entry = reviewed.find((candidate) => isReviewed(candidate, result));
		if (entry) {
			used.add(entry);
			accepted.push(result);
		} else {
			open.push(result);
		}
	}
	return { open, accepted, stale: reviewed.filter((entry) => !used.has(entry)) };
}

/**
 * Whether this scan's results cover only part of the code. On a pull request
 * CodeQL runs diff-informed: it reports only results in the changed lines
 * (the log's "Computing PR diff ranges"), so a reviewed entry elsewhere finds
 * no result and would look stale on every pull request that does not touch
 * it. Stale entries are judged by the push and scheduled runs on main, which
 * see everything. The malware scans always cover every file.
 */
export function partialResults(scan, env) {
	return scan.startsWith('CodeQL ') && ['pull_request', 'pull_request_target'].includes(env.GITHUB_EVENT_NAME);
}

export function describeEntry(entry) {
	return entry.file === undefined
		? `${entry.rule} at ${entry.path}: \`${entry.line}\``
		: `${entry.rule} at ${entry.file}`;
}

if (isMain(import.meta)) {
	const [scan, directory, reviewedFile] = process.argv.slice(2);
	try {
		if (!SCANS.includes(scan)) {
			throw new Error(`Unknown scan ${JSON.stringify(scan)}; expected one of ${SCANS.join(', ')}`);
		}
		const files = (fs.existsSync(directory) ? fs.readdirSync(directory) : [])
			.filter((name) => name.endsWith('.sarif'))
			.map((name) => path.join(directory, name));
		const { open, accepted, stale } = judge(sarifResults(files), loadReviewed(reviewedFile, scan));
		console.log(`${scan}: ${open.length} unreviewed result(s), ${accepted.length} reviewed, ${stale.length} stale entr${stale.length === 1 ? 'y' : 'ies'}`);
		for (const result of open) {
			console.log(`  ${result.rule}: ${result.uri}:${result.line ?? '?'} ${result.message}`);
		}
		for (const entry of stale) {
			console.log(`  Reviewed, but no longer found: ${describeEntry(entry)}`);
		}
		if (open.length > 0 || (stale.length > 0 && !partialResults(scan, process.env))) {
			process.exitCode = 1;
		} else if (stale.length > 0) {
			console.log('  Not failed: on a pull request CodeQL reports only the changed code, so an entry outside it'
				+ ' looks stale. The run on main fails on it.');
		}
	} catch (error) {
		console.error(error.message);
		process.exitCode = 1;
	}
}
