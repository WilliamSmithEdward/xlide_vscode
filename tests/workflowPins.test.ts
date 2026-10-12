// Every action a workflow uses is pinned to a full commit SHA whose comment
// names that commit's exact release, `# v4.38.2`. A major or minor tag in the
// comment (`# v4`) moves: once the upstream moves it, zizmor's
// ref-version-mismatch flags a pin that never changed, as it did for the
// codeql-action pins in security.yml (#1334). This catches the comment when it
// is written, not weeks later.

import { readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import { describe, expect, it } from 'vitest';

const ROOT = join(__dirname, '..', '.github');

/** Every workflow and composite-action file under .github. */
function workflowFiles(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) {
			return workflowFiles(path);
		}
		return /\.ya?ml$/.test(entry.name) ? [path] : [];
	});
}

/**
 * What is wrong with one `uses:` line, or undefined when nothing is. A local
 * action (`./...`) is pinned by the repository's own commit, and a container
 * (`docker://`) by its digest.
 */
function pinProblem(line: string): string | undefined {
	const match = /^\s*-?\s*uses:\s*(\S+)(.*)$/.exec(line);
	if (!match) {
		return undefined;
	}
	const [, target, rest] = match;
	if (target.startsWith('./') || target.startsWith('docker://')) {
		return undefined;
	}
	if (!/@[0-9a-f]{40}$/.test(target)) {
		return 'not pinned to a full commit SHA';
	}
	const comment = /^\s+#\s*(\S+)\s*$/.exec(rest)?.[1];
	if (!comment) {
		return 'no version comment';
	}
	return /^v?\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(comment) ? undefined : `comment '${comment}' is not an exact release tag`;
}

describe('workflow action pins', () => {
	it('pin every action to a SHA commented with its exact release', () => {
		const problems = workflowFiles(ROOT).flatMap((file) => readFileSync(file, 'utf8').split(/\r?\n/)
			.map((line, i) => ({ problem: pinProblem(line), at: `${file.slice(ROOT.length + 1)}:${i + 1}` }))
			.filter((hit) => hit.problem !== undefined)
			.map((hit) => `${hit.at}: ${hit.problem}`));
		expect(problems).toEqual([]);
	});

	it('refuse a moving tag, a missing comment and an unpinned ref, and let the rest through', () => {
		const sha = 'a'.repeat(40);
		expect(pinProblem(`      - uses: actions/checkout@${sha} # v7.0.1`)).toBeUndefined();
		expect(pinProblem(`        uses: github/codeql-action/init@${sha} # v4.38.2`)).toBeUndefined();
		expect(pinProblem(`      - uses: some/action@${sha} # 1.2.3`)).toBeUndefined();
		expect(pinProblem('      - uses: ./.github/actions/scan-targets')).toBeUndefined();
		expect(pinProblem('      - uses: docker://alpine@sha256:abc')).toBeUndefined();
		expect(pinProblem('      - run: echo uses: nothing')).toBeUndefined();
		expect(pinProblem(`      - uses: actions/checkout@${sha} # v4`)).toMatch(/not an exact release tag/);
		expect(pinProblem(`      - uses: actions/checkout@${sha} # v4.38`)).toMatch(/not an exact release tag/);
		expect(pinProblem(`      - uses: actions/checkout@${sha}`)).toMatch(/no version comment/);
		expect(pinProblem('      - uses: actions/checkout@v4')).toMatch(/full commit SHA/);
		expect(pinProblem('      - uses: actions/checkout@main # v7.0.1')).toMatch(/full commit SHA/);
	});
});
