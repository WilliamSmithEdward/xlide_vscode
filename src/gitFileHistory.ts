// A file's history in git: where its repository is, whether it is tracked,
// the commits that touched it, and its bytes at a revision.
//
// Everything goes through the `git` executable so the answer is the one the
// user's own git gives - the same repository roots, the same worktrees, the
// same credentials-free reads of local objects. No `vscode` dependency: the
// command layer chooses the executable and turns refusals into messages.

import { execFile } from 'node:child_process';
import * as path from 'node:path';

export interface GitCommit {
	hash: string;
	shortHash: string;
	author: string;
	/** ISO date (YYYY-MM-DD) of the author date. */
	date: string;
	subject: string;
}

/** A tracked file: its repository's working tree root and its path inside it. */
export interface GitFileRef {
	root: string;
	/** Forward-slashed path relative to the root, the way git names it. */
	relativePath: string;
}

export interface GitRunResult {
	code: number;
	stdout: Buffer;
	stderr: string;
}

export interface GitRunner {
	run(args: readonly string[], cwd: string): Promise<GitRunResult>;
}

/** Thrown when the git executable itself cannot be started. */
export class GitUnavailableError extends Error {
	constructor(readonly executable: string, cause: string) {
		super(`Could not run "${executable}": ${cause}`);
		this.name = 'GitUnavailableError';
	}
}

/** A workbook's git object can be tens of megabytes; give `git show` room. */
const MAX_OUTPUT_BYTES = 512 * 1024 * 1024;

/** Runs the given git executable (or the one on PATH). */
export function gitRunner(executable = 'git'): GitRunner {
	return {
		run(args, cwd) {
			return new Promise((resolve, reject) => {
				execFile(
					executable,
					[...args],
					{ cwd, encoding: 'buffer', maxBuffer: MAX_OUTPUT_BYTES, windowsHide: true },
					(error, stdout, stderr) => {
						const failure = error as (NodeJS.ErrnoException & { code?: number | string }) | null;
						if (failure && typeof failure.code === 'string') {
							// ENOENT and friends: the process never ran.
							reject(new GitUnavailableError(executable, failure.message));
							return;
						}
						resolve({
							code: failure ? (typeof failure.code === 'number' ? failure.code : 1) : 0,
							stdout: Buffer.isBuffer(stdout) ? stdout : Buffer.from(stdout ?? ''),
							stderr: Buffer.isBuffer(stderr) ? stderr.toString('utf8') : String(stderr ?? ''),
						});
					},
				);
			});
		},
	};
}

/**
 * The repository a file sits in, or undefined when no repository holds it.
 * Throws {@link GitUnavailableError} when git cannot run at all.
 */
export async function gitFileRef(filePath: string, git: GitRunner): Promise<GitFileRef | undefined> {
	const directory = path.dirname(filePath);
	const result = await git.run(['rev-parse', '--show-toplevel'], directory);
	if (result.code !== 0) {
		return undefined;
	}
	const root = path.normalize(result.stdout.toString('utf8').trim());
	if (!root) {
		return undefined;
	}
	const relativePath = path.relative(root, filePath).split(path.sep).join('/');
	if (relativePath.startsWith('..')) {
		return undefined;
	}
	return { root, relativePath };
}

/** Whether git tracks the file (a never-committed file has no history to compare). */
export async function gitFileIsTracked(ref: GitFileRef, git: GitRunner): Promise<boolean> {
	const result = await git.run(['ls-files', '--error-unmatch', '--', ref.relativePath], ref.root);
	return result.code === 0;
}

/** The commits that touched the file, newest first. */
export async function gitFileHistory(ref: GitFileRef, git: GitRunner, limit = 50): Promise<GitCommit[]> {
	const separator = '\u001f';
	const result = await git.run([
		'log',
		`--max-count=${limit}`,
		`--format=%H${separator}%h${separator}%an${separator}%as${separator}%s`,
		'--',
		ref.relativePath,
	], ref.root);
	if (result.code !== 0) {
		return [];
	}
	const commits: GitCommit[] = [];
	for (const line of result.stdout.toString('utf8').split(/\r?\n/)) {
		if (!line) {
			continue;
		}
		const [hash, shortHash, author, date, ...subject] = line.split(separator);
		if (!hash) {
			continue;
		}
		commits.push({ hash, shortHash, author, date, subject: subject.join(separator) });
	}
	return commits;
}

/**
 * The file's bytes as committed at `revision`, or undefined when that
 * revision has no such file (added since, or deleted before).
 */
export async function gitFileAtRevision(
	ref: GitFileRef,
	revision: string,
	git: GitRunner,
): Promise<Buffer | undefined> {
	const result = await git.run(['show', `${revision}:${ref.relativePath}`], ref.root);
	if (result.code !== 0) {
		return undefined;
	}
	return result.stdout;
}

/** The full hash a revision names in the file's repository, or undefined when it names nothing. */
export async function gitRevisionHash(ref: GitFileRef, revision: string, git: GitRunner): Promise<string | undefined> {
	const result = await git.run(['rev-parse', '--verify', `${revision}^{commit}`], ref.root);
	if (result.code !== 0) {
		return undefined;
	}
	const hash = result.stdout.toString('utf8').trim();
	return /^[0-9a-f]{40}$/i.test(hash) ? hash : undefined;
}

/** The short form of a revision for a diff title: `HEAD`, or `abc1234`. */
export function gitRevisionLabel(revision: string, commit?: GitCommit): string {
	if (commit) {
		return commit.shortHash;
	}
	return /^[0-9a-f]{40}$/i.test(revision) ? revision.slice(0, 7) : revision;
}
