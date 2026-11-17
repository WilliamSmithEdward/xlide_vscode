// Fails when npm audit reports an advisory that is not in the reviewed list,
// when a reviewed entry is past its review date, or when an entry matches no
// advisory: an entry left behind would accept the next one to land there.
//   node .github/scripts/npm-audit-gate.mjs <npm-audit.json> <package-lock.json> [reviewed.json]
// An entry names the advisory, the package and the installed version it was
// reviewed for, so a new version of the package comes back for review. A
// package npm lists only because it depends on a vulnerable one carries no
// advisory of its own and is judged where the advisory is.
import fs from 'node:fs';
import { isMain } from './scan-sarif.mjs';

const ADVISORY_ID = /\/(GHSA(?:-[23456789cfghjmpqrvwx]{4}){3})$/;

function nonEmpty(value) {
	return typeof value === 'string' && value.trim() !== '';
}

/** Each advisory npm audit reports, once per installed version of its package. */
export function auditAdvisories(audit, lockfile) {
	const found = [];
	for (const [name, entry] of Object.entries(audit.vulnerabilities ?? {})) {
		const versions = [...new Set((entry.nodes ?? []).map((node) => {
			const version = lockfile.packages?.[node]?.version;
			if (!version) {
				throw new Error(`${name}: ${node} is not in the audited lockfile`);
			}
			return version;
		}))];
		for (const via of entry.via ?? []) {
			if (typeof via === 'string') {
				continue;
			}
			const advisory = ADVISORY_ID.exec(via.url ?? '')?.[1];
			if (!advisory || via.name !== name) {
				throw new Error(`${name}: unreadable advisory ${JSON.stringify(via.url ?? via)}`);
			}
			for (const version of versions) {
				found.push({ advisory, package: name, version, severity: via.severity, title: via.title });
			}
		}
	}
	const total = audit.metadata?.vulnerabilities?.total ?? 0;
	if (total > 0 && found.length === 0) {
		throw new Error(`npm audit reports ${total} vulnerabilities but no advisory could be read`);
	}
	return found;
}

/** The reviewed list; a malformed entry refuses the whole list. */
export function loadReviewed(file) {
	if (!file || !fs.existsSync(file)) {
		return [];
	}
	const entries = JSON.parse(fs.readFileSync(file, 'utf8')).reviewed;
	if (!Array.isArray(entries)) {
		throw new Error(`${file}: "reviewed" is not a list`);
	}
	return entries.map((entry, index) => {
		const where = `${file}, entry ${index + 1}`;
		for (const field of ['advisory', 'package', 'version', 'reason', 'reviewBy']) {
			if (!nonEmpty(entry[field])) {
				throw new Error(`${where}: needs ${field}`);
			}
		}
		if (!/^\d{4}-\d\d-\d\d$/.test(entry.reviewBy) || Number.isNaN(Date.parse(`${entry.reviewBy}T23:59:59Z`))) {
			throw new Error(`${where}: reviewBy must be a YYYY-MM-DD date`);
		}
		return entry;
	});
}

/** Splits advisories into open and accepted, and finds stale or expired entries. */
export function judge(advisories, reviewed, now = Date.now()) {
	const matches = (entry, found) => entry.advisory === found.advisory
		&& entry.package === found.package && entry.version === found.version;
	const accepted = advisories.filter((found) => reviewed.some((entry) => matches(entry, found)));
	return {
		open: advisories.filter((found) => !accepted.includes(found)),
		accepted,
		stale: reviewed.filter((entry) => !advisories.some((found) => matches(entry, found))),
		expired: reviewed.filter((entry) => now > Date.parse(`${entry.reviewBy}T23:59:59Z`)),
	};
}

export const describe = (item) => `${item.advisory} in ${item.package} ${item.version}`;

if (isMain(import.meta)) {
	const [auditFile, lockFile, reviewedFile] = process.argv.slice(2);
	if (!auditFile || !lockFile) {
		console.error('usage: npm-audit-gate.mjs <npm-audit.json> <package-lock.json> [reviewed.json]');
		process.exit(2);
	}
	const read = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));
	const { open, accepted, stale, expired } = judge(auditAdvisories(read(auditFile), read(lockFile)), loadReviewed(reviewedFile));
	for (const found of accepted) {
		console.log(`Reviewed: ${describe(found)} (${found.severity}), see ${reviewedFile}`);
	}
	const problems = [
		...open.map((found) => `Not reviewed: ${describe(found)} (${found.severity}) ${found.title}`),
		...expired.map((entry) => `Past its review date ${entry.reviewBy}: ${describe(entry)}`),
		...stale.map((entry) => `Reviewed, but no longer reported: ${describe(entry)}`),
	];
	for (const problem of problems) {
		console.error(problem);
	}
	process.exit(problems.length ? 1 : 0);
}
