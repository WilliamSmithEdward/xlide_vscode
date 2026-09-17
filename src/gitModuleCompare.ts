// Compare a module, or a whole project, with what git committed.
//
// A workbook is one binary blob to git, so the Source Control view can only
// say that it changed. XLIDE parses the committed bytes the same way it
// parses the file on disk, so a module can be diffed against HEAD or any
// earlier commit without exporting anything, in the same diff editor the
// agent review uses. A VB6 project keeps its modules as files, so those read
// straight from the commit.
//
// Everything git-specific is in gitFileHistory.ts; this layer is the
// commands, the picker and the read-only document the diff's left side reads.

import * as vscode from 'vscode';
import * as path from 'node:path';
import type { ProjectEngine } from './projectEngine';
import type { XlideNode } from './projectExplorer';
import { registerXlideCommand } from './xlideCommandRegistration';
import { XLIDE_SCHEME, decodeModuleUri, encodeModuleUri } from './xlideFileSystem';
import { errorMessage } from './util/errors';
import { evictOldest } from './util/boundedMap';
import { resolveProjectPath, statusMessage } from './commands/shared';
import { isVb6ProjectPath } from './macroContainerUi';
import { readModulesFromBuffer, type ModuleEntry } from './vba/projectService';
import { decodeCodePage } from './vba/codePages';
import { VB6_CODE_PAGE } from './vba/vb6/vb6Project';
import {
	GitUnavailableError,
	gitFileAtRevision,
	gitFileHistory,
	gitFileIsTracked,
	gitFileRef,
	gitRevisionLabel,
	gitRunner,
	type GitCommit,
	type GitFileRef,
	type GitRunner,
} from './gitFileHistory';
import { unifiedDiff } from './util/lineDiff';

export const XLIDE_GIT_SCHEME = 'xlide-vba-git';

/** A module's name and the code the editor shows for it, at one revision. */
export interface ModuleSnapshot {
	name: string;
	source: string;
	/** The module's own file, for the containers whose modules are files. */
	filePath?: string;
}

export type ModuleChangeKind = 'modified' | 'added' | 'removed';

export interface ModuleChange {
	name: string;
	kind: ModuleChangeKind;
}

export interface GitModuleCompareDeps {
	git: GitRunner;
	/** Every module the project has on disk now, as the editor would show it. */
	currentModules(projectPath: string): Promise<ModuleSnapshot[]>;
	/** Every module the project had at a revision, or undefined when the project was not in it. */
	modulesAtRevision(projectPath: string, ref: GitFileRef, revision: string): Promise<ModuleSnapshot[] | undefined>;
}

/** Modules as git holds them, keyed by lowercased name. */
type Snapshots = Map<string, ModuleSnapshot>;

const committedSources = new Map<string, string>();
let versionCounter = 0;

/** Serves the committed module text the diff view's left side reads. */
export function registerGitCompareContentProvider(): vscode.Disposable {
	return vscode.workspace.registerTextDocumentContentProvider(XLIDE_GIT_SCHEME, {
		provideTextDocumentContent(uri: vscode.Uri): string {
			return committedSources.get(uri.toString()) ?? '';
		},
	});
}

/** The deps the extension uses: the user's git, and the engine's module reads. */
export function gitModuleCompareDeps(bridge: ProjectEngine): GitModuleCompareDeps {
	return {
		get git() {
			return gitRunner(configuredGitPath());
		},
		async currentModules(projectPath) {
			const modules = await bridge.call<ModuleEntry[]>('readModules', { path: projectPath });
			return modules
				.filter((module) => module.source !== undefined)
				.map((module) => ({ name: module.name, source: module.source ?? '', filePath: module.filePath }));
		},
		async modulesAtRevision(projectPath, ref, revision) {
			if (isVb6ProjectPath(projectPath)) {
				// The .vbp lists files; each one has its own committed bytes.
				const current = await this.currentModules(projectPath);
				const out: ModuleSnapshot[] = [];
				for (const module of current) {
					if (!module.filePath) {
						continue;
					}
					const fileRef = await gitFileRef(module.filePath, this.git);
					const bytes = fileRef ? await gitFileAtRevision(fileRef, revision, this.git) : undefined;
					if (bytes) {
						// A VB6 file is ANSI, and git keeps its bytes as they are.
						out.push({ name: module.name, source: decodeCodePage(bytes, VB6_CODE_PAGE), filePath: module.filePath });
					}
				}
				return out;
			}
			const bytes = await gitFileAtRevision(ref, revision, this.git);
			if (!bytes) {
				return undefined;
			}
			return readModulesFromBuffer(bytes)
				.filter((module) => module.source !== undefined)
				.map((module) => ({ name: module.name, source: module.source ?? '' }));
		},
	};
}

/** The git executable the user's git extension is configured with, or `git`. */
function configuredGitPath(): string {
	const configured = vscode.workspace.getConfiguration('git').get<string | string[] | null>('path');
	if (typeof configured === 'string' && configured.trim()) {
		return configured.trim();
	}
	if (Array.isArray(configured)) {
		const first = configured.find((entry) => typeof entry === 'string' && entry.trim());
		if (first) {
			return first.trim();
		}
	}
	return 'git';
}

export function registerGitCompareCommands(
	bridge: ProjectEngine,
	deps: GitModuleCompareDeps = gitModuleCompareDeps(bridge),
): vscode.Disposable[] {
	return [
		registerGitCompareContentProvider(),
		registerXlideCommand('xlide.compareModuleWithHead', async (node?: Partial<XlideNode>) => {
			const target = moduleTarget(node);
			if (!target) {
				return;
			}
			await compareModule(deps, target.projectPath, target.moduleName, 'HEAD');
		}),
		registerXlideCommand('xlide.compareModuleWithRevision', async (node?: Partial<XlideNode>) => {
			const target = moduleTarget(node);
			if (!target) {
				return;
			}
			const commit = await pickCommit(deps, target.projectPath);
			if (commit) {
				await compareModule(deps, target.projectPath, target.moduleName, commit.hash, commit);
			}
		}),
		registerXlideCommand('xlide.compareProjectWithHead', async (node?: Partial<XlideNode>) => {
			const projectPath = projectTarget(node);
			if (!projectPath) {
				return;
			}
			await compareProject(deps, projectPath, 'HEAD');
		}),
		registerXlideCommand('xlide.restoreModuleFromHead', async (node?: Partial<XlideNode>, options: RestoreModuleOptions = {}) => {
			const target = moduleTarget(node);
			if (!target) {
				return;
			}
			await restoreModule(deps, target.projectPath, target.moduleName, options);
		}),
		registerXlideCommand('xlide.moduleHistory', async (node?: Partial<XlideNode>, options: ModuleHistoryOptions = {}) => {
			const target = moduleTarget(node);
			if (!target) {
				return;
			}
			await moduleHistory(deps, target.projectPath, target.moduleName, options);
		}),
	];
}

// ---------------------------------------------------------------- history

export interface ModuleHistoryOptions {
	/** Open the diff of this commit (hash or prefix) without the picker; the integration suite has no one to pick. */
	commit?: string;
	/** How many of the file's commits to look through, newest first. */
	limit?: number;
}

export type ModuleHistoryKind = 'modified' | 'added' | 'removed';

/** One commit in which the module's text differed from the commit before it. */
export interface ModuleHistoryEntry {
	commit: GitCommit;
	kind: ModuleHistoryKind;
	/** The module in this commit; undefined when the commit removed it. */
	source: string | undefined;
	/** The module as it was before this commit; undefined when the commit added it. */
	previous: string | undefined;
	/** The commit the previous text comes from; undefined when this is the oldest commit looked at. */
	previousCommit: GitCommit | undefined;
}

const DEFAULT_HISTORY_LIMIT = 50;

/**
 * The commits that changed one module, out of the commits that touched the
 * file. `git log` can only say the file changed; this reads the module out
 * of every commit in the window and keeps the ones where its text moved.
 * Commits are compared each with the one before it in the file's history,
 * so a commit's entry is what that commit did to the module.
 */
export async function moduleHistoryEntries(
	deps: GitModuleCompareDeps,
	projectPath: string,
	ref: GitFileRef,
	moduleName: string,
	commits: readonly GitCommit[],
	options: {
		/** True when `commits` is the file's whole history, so its oldest commit is where the module first appeared. */
		windowComplete: boolean;
		progress?: (done: number, total: number, commit: GitCommit) => void;
		isCancelled?: () => boolean;
	},
): Promise<ModuleHistoryEntry[]> {
	const wanted = moduleName.toLowerCase();
	// Oldest first, so each commit's predecessor is already in hand.
	const ordered = [...commits].reverse();
	const texts: Array<string | undefined> = [];
	for (const [index, commit] of ordered.entries()) {
		if (options.isCancelled?.()) {
			return [];
		}
		options.progress?.(index + 1, ordered.length, commit);
		const modules = await snapshotsAt(deps, projectPath, ref, commit.hash);
		texts.push(modules?.get(wanted)?.source);
	}
	const entries: ModuleHistoryEntry[] = [];
	for (let i = 0; i < ordered.length; i++) {
		const source = texts[i];
		const previous = i > 0 ? texts[i - 1] : undefined;
		if (i === 0) {
			// The oldest commit read. When the window holds the whole history
			// it is where the module first appeared; when the window is cut
			// short, nothing older was read and the commit can say nothing.
			if (options.windowComplete && source !== undefined) {
				entries.push({ commit: ordered[i], kind: 'added', source, previous: undefined, previousCommit: undefined });
			}
			continue;
		}
		if (source === undefined && previous === undefined) {
			continue;
		}
		if (source === undefined) {
			entries.push({ commit: ordered[i], kind: 'removed', source, previous, previousCommit: ordered[i - 1] });
		} else if (previous === undefined) {
			entries.push({ commit: ordered[i], kind: 'added', source, previous, previousCommit: ordered[i - 1] });
		} else if (normalize(source) !== normalize(previous)) {
			entries.push({ commit: ordered[i], kind: 'modified', source, previous, previousCommit: ordered[i - 1] });
		}
	}
	return entries.reverse();
}

/**
 * The committed modules of a revision, read once per commit and kept for
 * the next question. A commit's content never changes, so an entry is valid
 * for as long as it is kept.
 */
const snapshotCache = new Map<string, Snapshots | undefined>();
const SNAPSHOT_CACHE_CAP = 64;

export function resetGitSnapshotCacheForTests(): void {
	snapshotCache.clear();
}

async function snapshotsAt(
	deps: GitModuleCompareDeps,
	projectPath: string,
	ref: GitFileRef,
	hash: string,
): Promise<Snapshots | undefined> {
	const key = `${ref.root}\u0000${ref.relativePath}\u0000${hash}`;
	if (snapshotCache.has(key)) {
		const hit = snapshotCache.get(key);
		snapshotCache.delete(key);
		snapshotCache.set(key, hit);
		return hit;
	}
	const modules = snapshotsByName(await deps.modulesAtRevision(projectPath, ref, hash));
	snapshotCache.set(key, modules);
	evictOldest(snapshotCache, SNAPSHOT_CACHE_CAP);
	return modules;
}

/** Lists the commits that changed the module and opens the diff of the one picked. */
export async function moduleHistory(
	deps: GitModuleCompareDeps,
	projectPath: string,
	moduleName: string,
	options: ModuleHistoryOptions = {},
): Promise<void> {
	const ref = await trackedRef(deps, projectPath);
	if (!ref) {
		return;
	}
	const name = path.basename(projectPath);
	const limit = options.limit ?? DEFAULT_HISTORY_LIMIT;
	const commits = await gitFileHistory(ref, deps.git, limit);
	if (commits.length === 0) {
		void vscode.window.showInformationMessage(`XLIDE: git has no commits for "${name}".`);
		return;
	}
	let entries: ModuleHistoryEntry[];
	try {
		entries = await vscode.window.withProgress(
			{
				location: vscode.ProgressLocation.Notification,
				title: `XLIDE: Reading ${moduleName} out of ${commits.length} commits of ${name}`,
				cancellable: true,
			},
			(progress, token) => moduleHistoryEntries(deps, projectPath, ref, moduleName, commits, {
				windowComplete: commits.length < limit,
				progress: (done, total, commit) => progress.report({ message: `${commit.shortHash} (${done}/${total})`, increment: 100 / total }),
				isCancelled: () => token.isCancellationRequested,
			}),
		);
	} catch (err) {
		void vscode.window.showErrorMessage(`XLIDE: Could not read the history of "${moduleName}": ${errorMessage(err)}`);
		return;
	}
	if (entries.length === 0) {
		void vscode.window.showInformationMessage(
			`XLIDE: "${moduleName}" did not change in the last ${commits.length} ${commits.length === 1 ? 'commit' : 'commits'} of ${name}.`,
		);
		return;
	}
	let picked: ModuleHistoryEntry | undefined;
	if (options.commit) {
		const prefix = options.commit.toLowerCase();
		picked = entries.find((entry) => entry.commit.hash.toLowerCase().startsWith(prefix));
		if (!picked) {
			void vscode.window.showInformationMessage(`XLIDE: commit ${options.commit} did not change "${moduleName}".`);
			return;
		}
	} else {
		const icons: Record<ModuleHistoryKind, string> = {
			modified: '$(diff-modified)',
			added: '$(diff-added)',
			removed: '$(diff-removed)',
		};
		const item = await vscode.window.showQuickPick(
			entries.map((entry) => ({
				label: `${icons[entry.kind]} ${entry.commit.shortHash}  ${entry.commit.subject}`,
				description: `${entry.commit.author}, ${entry.commit.date}`,
				detail: entry.kind,
				entry,
			})),
			{
				title: `${moduleName}: ${entries.length === 1 ? '1 commit' : `${entries.length} commits`} changed it, of the last ${commits.length} that touched ${name}`,
				placeHolder: 'Pick a commit to see what it changed in the module',
			},
		);
		picked = item?.entry;
	}
	if (!picked) {
		return;
	}
	await openHistoryDiff(projectPath, moduleName, picked);
}

/** Two committed versions of a module side by side: before the commit, and in it. */
async function openHistoryDiff(projectPath: string, moduleName: string, entry: ModuleHistoryEntry): Promise<void> {
	const livePath = encodeModuleUri(projectPath, moduleName).path;
	versionCounter += 1;
	const leftLabel = entry.previousCommit ? entry.previousCommit.shortHash : 'before';
	const rightLabel = entry.commit.shortHash;
	const leftUri = committedSourceUri(livePath, `rev=${leftLabel}`, entry.previous ?? '');
	const rightUri = committedSourceUri(livePath, `rev=${rightLabel}`, entry.source ?? '');
	evictOldest(committedSources, COMMITTED_SOURCE_CAP);
	const title = entry.kind === 'added'
		? `${moduleName}: added in ${rightLabel}`
		: entry.kind === 'removed'
			? `${moduleName}: removed in ${rightLabel}`
			: `${moduleName}: ${leftLabel} ↔ ${rightLabel}`;
	await vscode.commands.executeCommand('vscode.diff', leftUri, rightUri, title, { preview: true });
}

export interface RestoreModuleOptions {
	/** Skip the confirmation; the integration suite has no one to click it. */
	withoutPrompt?: boolean;
}

/**
 * Puts one module back to its committed text without touching the rest of
 * the workbook. The committed text is applied as an edit to the module's
 * own document, so Ctrl+Z undoes it, and the document is saved through the
 * usual path when it had no unsaved edits; when it had some, the restore
 * lands on top of them and the save is left to the user.
 */
export async function restoreModule(
	deps: GitModuleCompareDeps,
	projectPath: string,
	moduleName: string,
	options: RestoreModuleOptions = {},
): Promise<void> {
	const ref = await trackedRef(deps, projectPath);
	if (!ref) {
		return;
	}
	const name = path.basename(projectPath);
	let committed: Snapshots | undefined;
	try {
		committed = snapshotsByName(await deps.modulesAtRevision(projectPath, ref, 'HEAD'));
	} catch (err) {
		void vscode.window.showErrorMessage(`XLIDE: Could not read "${name}" at HEAD: ${errorMessage(err)}`);
		return;
	}
	if (!committed) {
		void vscode.window.showInformationMessage(`XLIDE: "${name}" is not in HEAD, so there is nothing to restore from.`);
		return;
	}
	const before = committed.get(moduleName.toLowerCase());
	if (!before) {
		void vscode.window.showInformationMessage(
			`XLIDE: "${moduleName}" is not in HEAD. It was added since the last commit; delete it from the tree if you want the committed state.`,
		);
		return;
	}
	const current = (await deps.currentModules(projectPath)).find((module) => module.name.toLowerCase() === moduleName.toLowerCase());
	const uri = current?.filePath ? vscode.Uri.file(current.filePath) : encodeModuleUri(projectPath, moduleName);
	const document = await vscode.workspace.openTextDocument(uri);
	if (normalize(document.getText()) === normalize(before.source)) {
		statusMessage(`XLIDE: ${moduleName} already matches HEAD`);
		return;
	}
	if (!options.withoutPrompt) {
		const choice = await vscode.window.showWarningMessage(
			`Restore "${moduleName}" from HEAD?`,
			{
				modal: true,
				detail: "The module's current text is replaced by the version in the last commit. Undo in the editor brings it back.",
			},
			'Restore',
		);
		if (choice !== 'Restore') {
			return;
		}
	}
	const eol = document.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
	const restored = before.source.replace(/\r\n?|\n/g, eol);
	const wasDirty = document.isDirty;
	const edit = new vscode.WorkspaceEdit();
	edit.replace(uri, new vscode.Range(document.positionAt(0), document.positionAt(document.getText().length)), restored);
	if (!(await vscode.workspace.applyEdit(edit))) {
		void vscode.window.showErrorMessage(`XLIDE: Could not restore "${moduleName}": the editor refused the edit.`);
		return;
	}
	await vscode.window.showTextDocument(document, { preview: true });
	if (wasDirty) {
		statusMessage(`XLIDE: Restored ${moduleName} from HEAD in the editor; save to write it into ${name}`);
		return;
	}
	if (!(await document.save())) {
		void vscode.window.showErrorMessage(`XLIDE: Restored "${moduleName}" in the editor, but it could not be saved.`);
		return;
	}
	statusMessage(`XLIDE: Restored ${moduleName} from HEAD`);
}

/** The module a command was invoked on: the tree row, else the active module editor. */
function moduleTarget(node?: Partial<XlideNode>): { projectPath: string; moduleName: string } | undefined {
	if (node?.filePath && node.moduleName) {
		return { projectPath: node.filePath, moduleName: node.moduleName };
	}
	const document = vscode.window.activeTextEditor?.document;
	if (document?.uri.scheme === XLIDE_SCHEME) {
		const decoded = decodeModuleUri(document.uri);
		return { projectPath: decoded.projectPath, moduleName: decoded.moduleName };
	}
	void vscode.window.showInformationMessage('XLIDE: Select a module in the XLIDE tree, or open one, to compare it with git.');
	return undefined;
}

function projectTarget(node?: Partial<XlideNode>): string | undefined {
	const projectPath = resolveProjectPath(node);
	if (!projectPath) {
		void vscode.window.showInformationMessage('XLIDE: Select a file in the XLIDE tree, or open one of its modules, to compare it with git.');
	}
	return projectPath;
}

/**
 * The repository ref for a project, after the checks every command shares:
 * git runs, the file is in a repository, and git tracks it. Each refusal
 * is told to the user, and undefined comes back.
 */
async function trackedRef(deps: GitModuleCompareDeps, projectPath: string): Promise<GitFileRef | undefined> {
	const name = path.basename(projectPath);
	try {
		const ref = await gitFileRef(projectPath, deps.git);
		if (!ref) {
			void vscode.window.showInformationMessage(`XLIDE: "${name}" is not inside a git repository.`);
			return undefined;
		}
		if (!(await gitFileIsTracked(ref, deps.git))) {
			void vscode.window.showInformationMessage(`XLIDE: "${name}" is not tracked by git yet, so there is no committed version to compare with.`);
			return undefined;
		}
		return ref;
	} catch (err) {
		if (err instanceof GitUnavailableError) {
			void vscode.window.showWarningMessage(`XLIDE: ${err.message}. Install git, or set "git.path" in your settings.`);
			return undefined;
		}
		void vscode.window.showErrorMessage(`XLIDE: git failed for "${name}": ${errorMessage(err)}`);
		return undefined;
	}
}

async function pickCommit(deps: GitModuleCompareDeps, projectPath: string): Promise<GitCommit | undefined> {
	const ref = await trackedRef(deps, projectPath);
	if (!ref) {
		return undefined;
	}
	const commits = await gitFileHistory(ref, deps.git);
	if (commits.length === 0) {
		void vscode.window.showInformationMessage(`XLIDE: git has no commits for "${path.basename(projectPath)}".`);
		return undefined;
	}
	const picked = await vscode.window.showQuickPick(
		commits.map((commit) => ({
			label: `$(git-commit) ${commit.shortHash}  ${commit.subject}`,
			description: `${commit.author}, ${commit.date}`,
			commit,
		})),
		{ title: `Compare ${path.basename(projectPath)} with a commit`, placeHolder: 'The commit to compare the current module with' },
	);
	return picked?.commit;
}

/** Opens the diff of one module between a revision and its current text. */
export async function compareModule(
	deps: GitModuleCompareDeps,
	projectPath: string,
	moduleName: string,
	revision: string,
	commit?: GitCommit,
): Promise<void> {
	const ref = await trackedRef(deps, projectPath);
	if (!ref) {
		return;
	}
	let committed: Snapshots | undefined;
	try {
		committed = snapshotsByName(await deps.modulesAtRevision(projectPath, ref, revision));
	} catch (err) {
		void vscode.window.showErrorMessage(
			`XLIDE: Could not read "${path.basename(projectPath)}" at ${gitRevisionLabel(revision, commit)}: ${errorMessage(err)}`,
		);
		return;
	}
	const label = gitRevisionLabel(revision, commit);
	if (!committed) {
		void vscode.window.showInformationMessage(
			`XLIDE: "${path.basename(projectPath)}" is not in ${label}, so there is nothing to compare with.`,
		);
		return;
	}
	const before = committed.get(moduleName.toLowerCase());
	const current = (await deps.currentModules(projectPath)).find((module) => module.name.toLowerCase() === moduleName.toLowerCase());
	await openDiff(projectPath, moduleName, label, before, current);
}

/**
 * Lists the modules that differ between a revision and the file on disk,
 * and opens the diff of the one picked.
 */
export async function compareProject(deps: GitModuleCompareDeps, projectPath: string, revision: string): Promise<void> {
	const ref = await trackedRef(deps, projectPath);
	if (!ref) {
		return;
	}
	const name = path.basename(projectPath);
	let committed: Snapshots | undefined;
	let current: Snapshots;
	try {
		committed = snapshotsByName(await deps.modulesAtRevision(projectPath, ref, revision));
		current = snapshotsByName(await deps.currentModules(projectPath)) ?? new Map();
	} catch (err) {
		void vscode.window.showErrorMessage(`XLIDE: Could not compare "${name}" with ${revision}: ${errorMessage(err)}`);
		return;
	}
	if (!committed) {
		void vscode.window.showInformationMessage(`XLIDE: "${name}" is not in ${revision}, so there is nothing to compare with.`);
		return;
	}
	const changes = moduleChanges(committed, current);
	if (changes.length === 0) {
		statusMessage(`XLIDE: No VBA changes in ${name} since ${revision}`);
		return;
	}
	const icons: Record<ModuleChangeKind, string> = {
		modified: '$(diff-modified)',
		added: '$(diff-added)',
		removed: '$(diff-removed)',
	};
	const picked = await vscode.window.showQuickPick(
		changes.map((change) => ({
			label: `${icons[change.kind]} ${change.name}`,
			description: change.kind,
			change,
		})),
		{
			title: `${name}: ${changes.length === 1 ? '1 module differs' : `${changes.length} modules differ`} from ${revision}`,
			placeHolder: 'Pick a module to see its diff',
		},
	);
	if (!picked) {
		return;
	}
	await openDiff(
		projectPath,
		picked.change.name,
		revision,
		committed.get(picked.change.name.toLowerCase()),
		current.get(picked.change.name.toLowerCase()),
	);
}

/** One changed module with the diff that shows the change, for an agent or a report. */
export interface ModuleChangeWithDiff extends ModuleChange {
	/** Unified diff, committed on the left, current on the right; empty for a removed module's text. */
	diff: string;
}

/** What an agent gets when it asks what changed in a file since a revision. */
export type GitChangesReport =
	| { filePath: string; revision: string; tracked: false; reason: string; changes: [] }
	| { filePath: string; revision: string; tracked: true; head: string; changes: ModuleChangeWithDiff[] };

/**
 * The changes between a revision and the file on disk, each with its diff.
 * Refusals (no repository, untracked, absent at the revision) come back as a
 * report with `tracked: false` and a reason, never as a throw, so an agent
 * can read the answer either way.
 */
export async function gitChangesReport(
	deps: GitModuleCompareDeps,
	projectPath: string,
	revision = 'HEAD',
	moduleName?: string,
): Promise<GitChangesReport> {
	const name = path.basename(projectPath);
	const refuse = (reason: string): GitChangesReport => ({ filePath: projectPath, revision, tracked: false, reason, changes: [] });
	const ref = await gitFileRef(projectPath, deps.git);
	if (!ref) {
		return refuse(`${name} is not inside a git repository.`);
	}
	if (!(await gitFileIsTracked(ref, deps.git))) {
		return refuse(`${name} is not tracked by git, so it has no committed version to compare with.`);
	}
	const committed = snapshotsByName(await deps.modulesAtRevision(projectPath, ref, revision));
	if (!committed) {
		return refuse(`${name} is not in ${revision}.`);
	}
	const current = snapshotsByName(await deps.currentModules(projectPath)) ?? new Map();
	const label = gitRevisionLabel(revision);
	const wanted = moduleName?.toLowerCase();
	const changes: ModuleChangeWithDiff[] = [];
	for (const change of moduleChanges(committed, current)) {
		if (wanted !== undefined && change.name.toLowerCase() !== wanted) {
			continue;
		}
		const before = committed.get(change.name.toLowerCase())?.source ?? '';
		const after = current.get(change.name.toLowerCase())?.source ?? '';
		changes.push({
			...change,
			diff: unifiedDiff(before, after, {
				beforeLabel: `${change.name} (${label})`,
				afterLabel: `${change.name} (current)`,
			}),
		});
	}
	return { filePath: projectPath, revision, tracked: true, head: label, changes };
}

/** Added, removed and modified modules between two snapshots, by name. */
export function moduleChanges(committed: Snapshots, current: Snapshots): ModuleChange[] {
	const changes: ModuleChange[] = [];
	for (const [key, module] of current) {
		const before = committed.get(key);
		if (!before) {
			changes.push({ name: module.name, kind: 'added' });
		} else if (normalize(before.source) !== normalize(module.source)) {
			changes.push({ name: module.name, kind: 'modified' });
		}
	}
	for (const [key, module] of committed) {
		if (!current.has(key)) {
			changes.push({ name: module.name, kind: 'removed' });
		}
	}
	return changes.sort((a, b) => a.name.localeCompare(b.name));
}

/** Snapshots keyed by lower-cased module name, the way the comparisons look them up. */
export function snapshotsByName(modules: readonly ModuleSnapshot[]): Snapshots;
export function snapshotsByName(modules: readonly ModuleSnapshot[] | undefined): Snapshots | undefined;
export function snapshotsByName(modules: readonly ModuleSnapshot[] | undefined): Snapshots | undefined {
	return modules && new Map(modules.map((module) => [module.name.toLowerCase(), module]));
}

/** The engine's own comparison: exact text, end-of-line normalized. */
function normalize(source: string): string {
	return source.replace(/\r\n?/g, '\n').trimEnd();
}

/**
 * Opens the diff: the committed text on the left, the live document on the
 * right. A module the revision lacks diffs against nothing; one the file
 * lost shows its committed text against nothing.
 */
async function openDiff(
	projectPath: string,
	moduleName: string,
	label: string,
	before: ModuleSnapshot | undefined,
	current: ModuleSnapshot | undefined,
): Promise<void> {
	const liveUri = current?.filePath ? vscode.Uri.file(current.filePath) : encodeModuleUri(projectPath, moduleName);
	versionCounter += 1;
	const leftUri = committedSourceUri(liveUri.path, `rev=${encodeURIComponent(label)}`, before?.source ?? '');
	const rightUri = current ? liveUri : committedSourceUri(liveUri.path, 'removed', '');
	evictOldest(committedSources, COMMITTED_SOURCE_CAP);
	const title = before && current
		? `${moduleName}: ${label} ↔ current`
		: before
			? `${moduleName}: ${label} ↔ removed`
			: `${moduleName}: not in ${label} ↔ current`;
	await vscode.commands.executeCommand('vscode.diff', leftUri, rightUri, title, { preview: true });
}

const COMMITTED_SOURCE_CAP = 32;

/** A URI the content provider serves `text` under, beside the live module's path. */
function committedSourceUri(livePath: string, query: string, text: string): vscode.Uri {
	const uri = vscode.Uri.from({ scheme: XLIDE_GIT_SCHEME, path: livePath, query: `${query}&v=${versionCounter}` });
	committedSources.set(uri.toString(), text);
	return uri;
}
