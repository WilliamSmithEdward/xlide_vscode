// Which modules changed since the last commit, for the tree to mark.
//
// The Explorer shows a workbook as one modified file. This keeps, per tracked
// project, the modules that differ from HEAD - the same comparison Compare
// File with Git HEAD makes - so a module row can carry an `M` or `A` badge
// and the file row a count. The answer is computed once per change and
// remembered: a project's marks are keyed on the workbook's modification
// time and size, and every project's marks are dropped when the repository's
// HEAD, index or refs change.
//
// `marksFor` never blocks a row draw. It answers from the cache and, when the
// cache is stale or empty, schedules one computation and fires
// `onDidChange` when that lands, which is when the tree redraws the rows.

import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import {
	gitFileIsTracked,
	gitFileRef,
	type GitFileRef,
	gitRevisionHash,
} from './gitFileHistory';
import { moduleChanges, snapshotsByName, type GitModuleCompareDeps, type ModuleChangeKind } from './gitModuleCompare';
import { startPerformanceTrace } from './performanceTrace';

export interface GitProjectMarks {
	/** Modules that differ from HEAD, keyed by lowercased module name. */
	byModule: ReadonlyMap<string, ModuleChangeKind>;
	/** Modules HEAD had that the file no longer has. */
	removed: number;
	/** The commit compared against, full hash. */
	head: string;
}

/** What the tree asks: the marks a project has, if anything is known yet. */
export interface GitChangeMarksSource {
	/**
	 * The project's marks, or undefined while nothing is known or when git has
	 * no say (no repository, not tracked). Asking schedules a computation when
	 * the cache is stale, and `onDidChange` fires with the project path once
	 * it lands.
	 */
	marksFor(projectPath: string): GitProjectMarks | undefined;
	onDidChange: vscode.Event<string>;
}

interface CacheEntry {
	/** The stat the marks were computed against; -1 for an unreadable file. */
	mtimeMs: number;
	size: number;
	marks: GitProjectMarks | undefined;
	/**
	 * The repository or the file moved on. The marks are still served, so
	 * the rows keep them until the fresh answer lands instead of blinking.
	 */
	stale: boolean;
}

export interface GitChangeMarksDeps extends GitModuleCompareDeps {
	/** The file's modification stamp, or undefined when it cannot be read. */
	stat?(projectPath: string): { mtimeMs: number; size: number } | undefined;
}

export class GitChangeMarks implements GitChangeMarksSource, vscode.Disposable {
	private readonly _emitter = new vscode.EventEmitter<string>();
	readonly onDidChange = this._emitter.event;
	private readonly _cache = new Map<string, CacheEntry>();
	private readonly _inFlight = new Map<string, Promise<void>>();
	/** Projects whose marks the tree has asked for; only these are recomputed on a git change. */
	private readonly _wanted = new Set<string>();
	private _disposed = false;

	constructor(
		private readonly _deps: GitChangeMarksDeps,
		private readonly _log: (line: string) => void = () => undefined,
	) {}

	marksFor(projectPath: string): GitProjectMarks | undefined {
		this._wanted.add(projectPath);
		const entry = this._cache.get(projectPath);
		const stamp = this._stat(projectPath);
		if (entry && !entry.stale && stamp && entry.mtimeMs === stamp.mtimeMs && entry.size === stamp.size) {
			return entry.marks;
		}
		this._schedule(projectPath);
		return entry?.marks;
	}

	/** A project's file changed: its marks are recomputed on the next ask, or now if the tree has asked before. */
	invalidate(projectPath: string): void {
		const entry = this._cache.get(projectPath);
		if (entry) {
			entry.stale = true;
		}
		if (this._wanted.has(projectPath)) {
			this._schedule(projectPath);
		}
	}

	/** The repository changed (a commit, a checkout): every project's marks are recomputed. */
	invalidateAll(): void {
		for (const entry of this._cache.values()) {
			entry.stale = true;
		}
		for (const projectPath of this._wanted) {
			this._schedule(projectPath);
		}
	}

	dispose(): void {
		this._disposed = true;
		this._emitter.dispose();
	}

	private _stat(projectPath: string): { mtimeMs: number; size: number } | undefined {
		if (this._deps.stat) {
			return this._deps.stat(projectPath);
		}
		try {
			const stat = fs.statSync(projectPath);
			return { mtimeMs: stat.mtimeMs, size: stat.size };
		} catch {
			return undefined;
		}
	}

	private _schedule(projectPath: string): void {
		if (this._inFlight.has(projectPath)) {
			return;
		}
		const run = this._compute(projectPath)
			.catch((err) => {
				this._log(`[git marks] ${path.basename(projectPath)}: ${err instanceof Error ? err.message : String(err)}`);
			})
			.finally(() => {
				this._inFlight.delete(projectPath);
			});
		this._inFlight.set(projectPath, run);
	}

	private async _compute(projectPath: string): Promise<void> {
		const stamp = this._stat(projectPath);
		if (!stamp) {
			this._store(projectPath, { mtimeMs: -1, size: -1, marks: undefined, stale: false });
			return;
		}
		const trace = startPerformanceTrace('gitMarks', path.basename(projectPath));
		try {
			const marks = await this._marks(projectPath);
			this._store(projectPath, { ...stamp, marks, stale: false });
			trace.end('ok', path.basename(projectPath));
		} catch (err) {
			trace.end('failed', path.basename(projectPath));
			this._store(projectPath, { ...stamp, marks: undefined, stale: false });
			throw err;
		}
	}

	private async _marks(projectPath: string): Promise<GitProjectMarks | undefined> {
		const ref: GitFileRef | undefined = await gitFileRef(projectPath, this._deps.git);
		if (!ref || !(await gitFileIsTracked(ref, this._deps.git))) {
			return undefined;
		}
		const head = await gitRevisionHash(ref, 'HEAD', this._deps.git);
		if (!head) {
			return undefined;
		}
		const committed = await this._deps.modulesAtRevision(projectPath, ref, head);
		if (!committed) {
			return undefined;
		}
		const current = await this._deps.currentModules(projectPath);
		const changes = moduleChanges(snapshotsByName(committed), snapshotsByName(current));
		const byModule = new Map<string, ModuleChangeKind>();
		let removed = 0;
		for (const change of changes) {
			if (change.kind === 'removed') {
				removed++;
			} else {
				byModule.set(change.name.toLowerCase(), change.kind);
			}
		}
		return { byModule, removed, head };
	}

	private _store(projectPath: string, entry: CacheEntry): void {
		if (this._disposed) {
			return;
		}
		const previous = this._cache.get(projectPath);
		this._cache.set(projectPath, entry);
		if (!sameMarks(previous?.marks, entry.marks)) {
			this._emitter.fire(projectPath);
		}
	}
}

function sameMarks(a: GitProjectMarks | undefined, b: GitProjectMarks | undefined): boolean {
	if (!a || !b) {
		return a === b;
	}
	if (a.head !== b.head || a.removed !== b.removed || a.byModule.size !== b.byModule.size) {
		return false;
	}
	for (const [name, kind] of a.byModule) {
		if (b.byModule.get(name) !== kind) {
			return false;
		}
	}
	return true;
}

/**
 * Drops every project's marks when the repository behind the workspace
 * changes: a commit, a checkout, a reset. HEAD and the index move on all of
 * them. The workspace's own `.git` is watched by path; a repository above the
 * workspace is only seen through VS Code's git extension when that is on.
 */
export function watchRepositoriesForMarks(marks: GitChangeMarks): vscode.Disposable[] {
	const disposables: vscode.Disposable[] = [];
	const watcher = vscode.workspace.createFileSystemWatcher('**/.git/{HEAD,index,packed-refs}');
	const changed = (): void => marks.invalidateAll();
	disposables.push(watcher, watcher.onDidChange(changed), watcher.onDidCreate(changed), watcher.onDidDelete(changed));
	const gitExtension = vscode.extensions.getExtension<{ getAPI(version: 1): GitApiSubset }>('vscode.git');
	if (gitExtension?.isActive) {
		try {
			const api = gitExtension.exports.getAPI(1);
			const follow = (repository: GitRepositorySubset): void => {
				disposables.push(repository.state.onDidChange(changed));
			};
			api.repositories.forEach(follow);
			disposables.push(api.onDidOpenRepository(follow));
		} catch {
			// The path watcher above still covers the workspace's own repository.
		}
	}
	return disposables;
}

interface GitRepositorySubset {
	state: { onDidChange: vscode.Event<void> };
}

interface GitApiSubset {
	repositories: GitRepositorySubset[];
	onDidOpenRepository: vscode.Event<GitRepositorySubset>;
}
