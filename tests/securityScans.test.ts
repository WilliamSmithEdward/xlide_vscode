import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { judgeClamscan, readClamscan, readVersion } from '../.github/scripts/clamav-sarif.mjs';
import { CANARY_RULE, EICAR_HEX, toSarif } from '../.github/scripts/scan-sarif.mjs';
import {
	compileErrorCount,
	countRules,
	judgeYaraX,
	readMatches,
	scanErrors,
} from '../.github/scripts/yara-x-sarif.mjs';

// The ClamAV and YARA-X jobs of the Security workflow: what their converters
// read from each scanner, what they refuse to trust, and that the SARIF they
// write passes through the existing gate and release report. The canary is
// checked in memory only: written to disk here, it would be the EICAR test
// file this machine's antivirus reacts to.

const CANARY = 'scan-canary/eicar.com';
const repoRoot = path.join(__dirname, '..');
const temporary: string[] = [];

function tempDir(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-scans-'));
	temporary.push(dir);
	return dir;
}

afterEach(() => {
	for (const dir of temporary.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function clamscanOutput(found: string[], summary: Record<string, number>, extra: string[] = []): string {
	return [
		...found,
		...extra,
		'',
		'----------- SCAN SUMMARY -----------',
		`Known viruses: ${summary.known ?? 8712345}`,
		'Engine version: 1.5.4',
		'Scanned directories: 0',
		`Scanned files: ${summary.scanned ?? 31000}`,
		`Infected files: ${summary.infected ?? found.length}`,
		...(summary.errors ? [`Total errors: ${summary.errors}`] : []),
		'Data scanned: 812.33 MiB',
		'Time: 190.214 sec (3 m 10 s)',
	].join('\n');
}

describe('the canary', () => {
	it('is the EICAR test file, byte for byte', () => {
		const bytes = Buffer.from(EICAR_HEX, 'hex');
		expect(bytes.length).toBe(68);
		expect(createHash('md5').update(bytes).digest('hex')).toBe('44d88612fea8a8f36de82e1278abb02f');
		expect(createHash('sha256').update(bytes).digest('hex'))
			.toBe('275a021bbfb6489e54d471899f7db9d1663fc695ec2fe2a2c4538aabf651fd0f');
	});

	it('is what the canary rule looks for, at the start of the file', () => {
		const rule = fs.readFileSync(path.join(repoRoot, '.github/scans/canary.yar'), 'utf8');
		const body = rule.slice(rule.indexOf('{', rule.indexOf('$eicar')) + 1, rule.indexOf('}', rule.indexOf('$eicar')));
		expect(body.replace(/\s+/g, '').toUpperCase()).toBe(EICAR_HEX.toUpperCase());
		expect(rule).toContain(`rule ${CANARY_RULE}`);
		expect(rule).toContain('$eicar at 0');
	});
});

describe('the ClamAV converter', () => {
	it('reads detections, including a file name with spaces and a colon', () => {
		const scan = readClamscan(clamscanOutput([
			`${CANARY}: Win.Test.EICAR_HDB-1 FOUND`,
			'tests/fixtures/odd name: v2.xlsm: Doc.Dropper.Agent-1234 FOUND',
		], {}));
		expect(scan.detections).toEqual([
			{ file: CANARY, signature: 'Win.Test.EICAR_HDB-1' },
			{ file: 'tests/fixtures/odd name: v2.xlsm', signature: 'Doc.Dropper.Agent-1234' },
		]);
		expect(scan.summary['Scanned files']).toBe('31000');
		expect(scan.problems).toEqual([]);
	});

	it('gates on every detection but the canary, and trusts a run that exits 1', () => {
		const scan = readClamscan(clamscanOutput([
			`${CANARY}: Win.Test.EICAR_HDB-1 FOUND`,
			'node_modules/pkg/dist/index.js: Js.Malware.Agent-1 FOUND',
		], {}));
		expect(judgeClamscan(scan, 1, CANARY)).toEqual({
			findings: [{
				ruleId: 'Js.Malware.Agent-1',
				uri: 'node_modules/pkg/dist/index.js',
				message: 'ClamAV reports Js.Malware.Agent-1',
			}],
			problems: [],
		});
	});

	it('does not trust a run that missed the canary, exited otherwise, or counted errors', () => {
		const missed = readClamscan(clamscanOutput([], {}));
		expect(judgeClamscan(missed, 0, CANARY).problems).toEqual([
			'clamscan exited with 0; a sound run exits 1, having detected the canary',
			`clamscan did not detect the canary, ${CANARY}`,
		]);

		const errored = readClamscan(clamscanOutput([`${CANARY}: Win.Test.EICAR_HDB-1 FOUND`], { errors: 2 }, [
			'package/xlide.vsix: Can\'t allocate memory ERROR',
			'WARNING: tests/fixtures/gone.xlsm: Can\'t access file',
		]));
		expect(judgeClamscan(errored, 2, CANARY).problems).toEqual([
			'package/xlide.vsix: Can\'t allocate memory ERROR',
			'WARNING: tests/fixtures/gone.xlsm: Can\'t access file',
			'clamscan exited with 2; a sound run exits 1, having detected the canary',
			'clamscan counted 2 error(s)',
		]);
	});

	it('does not trust counts that disagree with what was printed', () => {
		const scan = readClamscan(clamscanOutput([`${CANARY}: Win.Test.EICAR_HDB-1 FOUND`], { infected: 3 }));
		expect(judgeClamscan(scan, 1, CANARY).problems).toEqual([
			'clamscan counted 3 infected file(s) but printed 1',
		]);
	});

	it('reads the engine and the daily signatures from the version line', () => {
		expect(readVersion('ClamAV 1.5.4/27791/Sun Sep 28 07:35:12 2026\n')).toEqual({
			engine: '1.5.4',
			database: '27791',
			databaseDate: 'Sun Sep 28 07:35:12 2026',
		});
	});
});

describe('the YARA-X converter', () => {
	const canaryMatch = { rule: CANARY_RULE, file: CANARY };

	it('reads the matches yr scan prints as JSON, with the rule description', () => {
		expect(readMatches({
			version: '1.20.0',
			matches: [
				{ rule: 'SIGNATURE_BASE_Webshell', file: 'package/contents/extension/out/extension.js', meta: { description: 'A web shell' } },
				{ rule: 'RULE_WITHOUT_META', file: 'README.md' },
			],
		})).toEqual([
			{ rule: 'SIGNATURE_BASE_Webshell', file: 'package/contents/extension/out/extension.js', description: 'A web shell' },
			{ rule: 'RULE_WITHOUT_META', file: 'README.md', description: undefined },
		]);
	});

	it('tells a file it could not scan from a rule it could not compile, colours or not', () => {
		const log = [
			'\u001b[1;31merror[E009]: \u001b[0munknown module `magic`',
			' --> yara-rules/packages/core/yara-rules-core.yar:1234:8',
			'error[E002]: syntax error',
			'\u001b[1;31merror: \u001b[0mscanning "node_modules/x/y.bin": Permission denied (os error 13)',
		].join('\n');
		expect(scanErrors(log)).toEqual(['error: scanning "node_modules/x/y.bin": Permission denied (os error 13)']);
		expect(compileErrorCount(log)).toBe(2);
	});

	it('counts the rules a source declares, private and global ones included', () => {
		expect(countRules([
			'import "pe"',
			'rule First : tag1 tag2 {',
			'    condition: true',
			'}',
			'private rule Second { condition: false }',
			'global private rule Third',
			'{ condition: true }',
			'  rule Fourth { strings: $a = "rule NotARule" condition: $a }',
		].join('\n'))).toBe(4);
	});

	it('gates on every match but the canary', () => {
		const verdict = judgeYaraX({
			matches: [canaryMatch, { rule: 'ELASTIC_Generic_Match', file: CANARY }, { rule: 'Real', file: './src/a.ts', description: 'Bad' }],
			scanLog: '',
			rulesLoaded: 950,
			rulesInSource: 1000,
			canary: CANARY,
			root: '/home/runner/work/xlide_vscode/xlide_vscode',
		});
		expect(verdict).toEqual({
			findings: [{ ruleId: 'Real', uri: 'src/a.ts', message: 'Real: Bad' }],
			problems: [],
		});
	});

	// The first run on GitHub: yr prints the absolute path of every file it
	// reads from a list, so the canary went unrecognised and was reported
	// twice, by the canary rule and by YARA Forge's own EICAR rule.
	it('finds the canary and the findings by workspace path, as yr prints them absolute', () => {
		const root = '/home/runner/work/xlide_vscode/xlide_vscode';
		const verdict = judgeYaraX({
			matches: [
				{ rule: 'TRELLIX_ARC_Malw_Eicar', file: `${root}/${CANARY}`, description: 'Rule to detect the EICAR pattern' },
				{ rule: CANARY_RULE, file: `${root}/${CANARY}` },
				{ rule: 'Real', file: `${root}/node_modules/pkg/index.js` },
			],
			scanLog: '',
			rulesLoaded: 5106,
			rulesInSource: 5111,
			canary: CANARY,
			root,
		});
		expect(verdict).toEqual({
			findings: [{ ruleId: 'Real', uri: 'node_modules/pkg/index.js', message: 'Real' }],
			problems: [],
		});
	});

	it('does not trust a run that missed the canary, failed on a file, or loaded too few rules', () => {
		const verdict = judgeYaraX({
			matches: [],
			scanLog: 'error: scanning "a": broken',
			rulesLoaded: 899,
			rulesInSource: 1000,
			canary: CANARY,
			root: '/home/runner/work/xlide_vscode/xlide_vscode',
		});
		expect(verdict.problems).toEqual([
			'error: scanning "a": broken',
			'YARA-X loaded 899 of 1000 rules, fewer than 90%',
			`YARA-X did not match the canary, ${CANARY}`,
		]);
	});
});

describe('the scan SARIF, through the existing gate and report', () => {
	function writeSarif(dir: string, name: string, driver: object, findings: object[], properties: object): void {
		fs.mkdirSync(dir, { recursive: true });
		fs.writeFileSync(path.join(dir, name), JSON.stringify(toSarif(driver, findings, properties)));
	}

	function gate(dir: string, reviewed: object): { status: number | null; output: string } {
		const reviewedFile = path.join(dir, 'reviewed.json');
		fs.writeFileSync(reviewedFile, JSON.stringify(reviewed));
		const run = spawnSync(process.execPath, ['.github/scripts/sarif-gate.mjs', 'Scan', dir, reviewedFile], {
			cwd: repoRoot,
			encoding: 'utf8',
		});
		return { status: run.status, output: run.stdout };
	}

	it('fails on a detection until it is reviewed, by signature and file', () => {
		const dir = tempDir();
		writeSarif(dir, 'clamav.sarif', { name: 'ClamAV', version: '1.5.4' }, [
			{ ruleId: 'Doc.Macro.Suspicious-1', uri: 'tests/fixtures/binaries/Book.xlsm', message: 'ClamAV reports Doc.Macro.Suspicious-1' },
		], {});

		const open = gate(dir, { reviewed: [] });
		expect(open.status).toBe(1);
		expect(open.output).toContain('Doc.Macro.Suspicious-1: tests/fixtures/binaries/Book.xlsm');

		const reviewed = gate(dir, {
			reviewed: [{ rule: 'Doc.Macro.Suspicious-1', file: 'tests/fixtures/binaries/*.xlsm', reason: 'A test fixture' }],
		});
		expect(reviewed.status).toBe(0);
	});

	it('passes a clean scan', () => {
		const dir = tempDir();
		writeSarif(dir, 'yara-x.sarif', { name: 'YARA-X', version: '1.20.0' }, [], {});
		expect(gate(dir, { reviewed: [] }).status).toBe(0);
	});

	it('puts the scanners, what they scanned and their detections in the release report', () => {
		const reports = tempDir();
		writeSarif(path.join(reports, 'clamav'), 'clamav.sarif', { name: 'ClamAV', version: '1.5.4' }, [], {
			engine: '1.5.4',
			database: '27791',
			databaseDate: 'Sun Sep 28 07:35:12 2026',
			knownSignatures: 8712345,
			scannedFiles: 30990,
			listedFiles: 31245,
		});
		writeSarif(path.join(reports, 'yara-x'), 'yara-x.sarif', { name: 'YARA-X', version: '1.20.0' }, [
			{ ruleId: 'SUSP_Generic', uri: 'node_modules/pkg/index.js', message: 'SUSP_Generic: Something' },
		], {
			engine: '1.20.0',
			rules: 'YARA Forge core 20260927',
			rulesLoaded: 1234,
			rulesInSource: 1300,
			listedFiles: 31245,
		});
		fs.mkdirSync(path.join(reports, 'package-sha256'));
		fs.writeFileSync(path.join(reports, 'package-sha256', 'package-sha256.txt'), `${'ab'.repeat(32)}  package/xlide-10.14.4.vsix\n`);
		const out = path.join(reports, 'malware-scan-report.md');
		const run = spawnSync(process.execPath, ['.github/scripts/malware-scan-report.mjs', reports, out], {
			cwd: repoRoot,
			encoding: 'utf8',
			env: { ...process.env, TAG: 'v10.14.4', CLAMAV_RESULT: 'success', YARA_X_RESULT: 'failure', PACKAGE_RESULT: 'success' },
		});
		expect(run.status).toBe(0);
		const report = fs.readFileSync(out, 'utf8');
		expect(report).toContain('# Malware scan report for v10.14.4');
		expect(report).toContain('| ClamAV | passed | 0 |');
		expect(report).toContain('| YARA-X (YARA Forge core 20260927) | did not pass (failure) | 1 |');
		expect(report).toContain(`\`xlide-10.14.4.vsix\`, the file attached to this release (SHA-256 \`${'ab'.repeat(32)}\`)`);
		expect(report).toContain('lockfile: 31,245 files.');
		expect(report).toContain('- ClamAV 1.5.4 with daily signatures 27791 (Sun Sep 28 07:35:12 2026), 8,712,345 signatures in all.');
		expect(report).toContain('- YARA-X 1.20.0 with YARA Forge core 20260927: 1,234 of 1,300 rules loaded.');
		expect(report).toContain('## YARA-X\n\n- `SUSP_Generic` at node_modules/pkg/index.js\n');
		expect(report).toContain('## ClamAV\n\nNo detections.\n');
	});
});

describe('the pinned YARA Forge release', () => {
	// The YARA Forge update workflow rewrites this file in each pull request
	// it opens; the Malware scan workflow reads the three fields back.
	it('names a release, its core package and that package s SHA-256', () => {
		const pin = JSON.parse(fs.readFileSync(path.join(repoRoot, '.github/scans/yara-forge.json'), 'utf8'));
		expect(pin.release).toMatch(/^\d{8}$/);
		expect(pin.asset).toBe('yara-forge-rules-core.zip');
		expect(pin.sha256).toMatch(/^[0-9a-f]{64}$/);
	});
});
