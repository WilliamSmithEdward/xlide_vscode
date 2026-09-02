// Review surface for AI-agent module writes.
//
// When Copilot (or any agent) writes VBA through xlide_writeModule, the edit
// never passes through the editor's edit pipeline, so Copilot's own
// keep/undo review cannot see it and no diff appears anywhere. VS Code's
// chat review tracks only the built-in file-edit tools, and no public API
// lets an extension tool's edits join it, so XLIDE supplies the same
// contract from its side using native surfaces only - no notifications:
// the write's before-image is kept, a diff opens quietly beside the chat
// (before <-> the live module document), and the XLIDE tree badges the
// module with inline Keep / Revert actions until the user decides. Revert
// restores the before-image through the audited write path - or removes
// the module when the agent's write created it. Later writes through any
// XLIDE path (editor saves included - Copilot's document edits arrive that
// way) keep the pending review tracking the live content, so Revert stays
// offered; only content that drifted outside every XLIDE write path makes
// it refuse.

import * as vscode from 'vscode';
import { encodeModuleUri, moduleIdentityKey, projectIdentityKey } from './xlideFileSystem';
import { errorMessage } from './util/errors';

export const XLIDE_AGENT_BEFORE_SCHEME = 'xlide-vba-before';

export interface AgentWriteRecord {
    before: string;
    /** False when the agent's write created the module; revert then deletes it. */
    beforeExisted: boolean;
    after: string;
}

const beforeImages = new Map<string, string>();
let writeCounter = 0;

// ---------------------------------------------------------------- pending

/** Agent writes awaiting a Keep/Revert decision, keyed by module identity.
 * The XLIDE tree badges these until the user resolves them. */
const pendingReviews = new Map<string, AgentWriteRecord>();
const pendingEmitter = new vscode.EventEmitter<{ filePath: string; moduleName: string }>();

/** Fires when a module gains or loses a pending agent review. */
export const onDidChangePendingAgentReviews = pendingEmitter.event;

function pendingKey(filePath: string, moduleName: string): string {
    return `${projectIdentityKey(filePath)}::${moduleIdentityKey(moduleName)}`;
}

export function hasPendingAgentReview(filePath: string, moduleName: string): boolean {
    return pendingReviews.has(pendingKey(filePath, moduleName));
}

function resolvePendingAgentReview(filePath: string, moduleName: string): void {
    if (pendingReviews.delete(pendingKey(filePath, moduleName))) {
        pendingEmitter.fire({ filePath, moduleName });
    }
}

/** Serves the frozen before-images the diff view's left side reads. */
export function registerAgentDiffProvider(): vscode.Disposable {
    return vscode.workspace.registerTextDocumentContentProvider(XLIDE_AGENT_BEFORE_SCHEME, {
        provideTextDocumentContent(uri: vscode.Uri): string {
            return beforeImages.get(uri.toString()) ?? '';
        },
    });
}

export function agentWriteDiffsEnabled(): boolean {
    return vscode.workspace
        .getConfiguration('xlide')
        .get<boolean>('agent.showWriteDiffs', true) === true;
}

export interface AgentWriteReviewDeps {
    /** Reads the module's current full source. */
    readModuleSource(filePath: string, moduleName: string): Promise<string>;
    /** Writes module source through the normal audited path. */
    writeModuleSource(filePath: string, moduleName: string, source: string): Promise<void>;
    /** Deletes the module through the normal audited path. */
    deleteModule(filePath: string, moduleName: string): Promise<void>;
}

/**
 * Registers an agent write for review and opens its diff. Stacked writes to
 * the same module merge: the review always compares against the state before
 * the agent's FIRST unreviewed write, so a revert restores what the user
 * last had, not the agent's own intermediate version.
 */
export async function presentAgentModuleWrite(
    filePath: string,
    moduleName: string,
    record: AgentWriteRecord,
): Promise<void> {
    const key = pendingKey(filePath, moduleName);
    const existing = pendingReviews.get(key);
    const merged: AgentWriteRecord = existing
        ? { before: existing.before, beforeExisted: existing.beforeExisted, after: record.after }
        : record;
    if (
        merged.beforeExisted &&
        normalizeForCompare(merged.before) === normalizeForCompare(merged.after)
    ) {
        // The write landed back on the pre-agent original - the agent undid
        // itself, or rewrote identical content. Nothing is left to review, and
        // a no-op diff would only mislead.
        resolvePendingAgentReview(filePath, moduleName);
        return;
    }
    pendingReviews.set(key, merged);
    pendingEmitter.fire({ filePath, moduleName });
    await openAgentReviewDiff(filePath, moduleName);
}

/**
 * Records a module write that did not present its own review: an editor save
 * (Copilot's document edits arrive that way too), a module-sync import, any
 * programmatic write. A pending review keeps tracking the live module - its
 * after-image follows the new content, so Revert stays offered and still
 * restores the pre-agent original the diff shows it discarding. Content that
 * lands back on the before-image resolves the review: nothing left to decide.
 */
export function trackModuleWriteForAgentReview(
    filePath: string,
    moduleName: string,
    newSource: string,
): void {
    const key = pendingKey(filePath, moduleName);
    const record = pendingReviews.get(key);
    if (!record) {
        return;
    }
    if (
        record.beforeExisted &&
        normalizeForCompare(newSource) === normalizeForCompare(record.before)
    ) {
        resolvePendingAgentReview(filePath, moduleName);
        return;
    }
    pendingReviews.set(key, { ...record, after: newSource });
}

/**
 * Opens the before/current diff for a module's pending agent write - also
 * the tree's "Review Agent Change" action. No-op when nothing is pending.
 */
export async function openAgentReviewDiff(filePath: string, moduleName: string): Promise<void> {
    const record = pendingReviews.get(pendingKey(filePath, moduleName));
    if (!record) {
        return;
    }
    const liveUri = encodeModuleUri(filePath, moduleName);
    // A fresh URI per view: an older diff tab keeps showing its own frozen
    // before-image instead of silently changing under the reader. The
    // counter rides the query so the path keeps its .bas extension - the
    // language association reads the path, and a suffixed extension would
    // render the before pane without VBA highlighting.
    writeCounter += 1;
    const beforeUri = vscode.Uri.from({
        scheme: XLIDE_AGENT_BEFORE_SCHEME,
        path: liveUri.path,
        query: `v${writeCounter}`,
    });
    beforeImages.set(beforeUri.toString(), record.before);
    pruneBeforeImages();
    try {
        await vscode.commands.executeCommand(
            'vscode.diff',
            beforeUri,
            liveUri,
            `${moduleName}: before agent edit ↔ current`,
            { preview: true, preserveFocus: true },
        );
    } catch {
        // Best-effort: the tree badge still holds the review.
    }
}

/** The tree's Keep action: the change stands, the badge clears. */
export function keepAgentChange(filePath: string, moduleName: string): void {
    resolvePendingAgentReview(filePath, moduleName);
}

/**
 * The tree's Revert action: restores the pre-agent state - the before-image
 * for an edited module, deletion for a module the agent created - unless the
 * module changed again after the agent's write.
 */
export async function revertAgentChange(
    deps: AgentWriteReviewDeps,
    filePath: string,
    moduleName: string,
): Promise<void> {
    const record = pendingReviews.get(pendingKey(filePath, moduleName));
    if (!record) {
        return;
    }
    try {
        const current = await deps.readModuleSource(filePath, moduleName);
        if (normalizeForCompare(current) !== normalizeForCompare(record.after)) {
            void vscode.window.showWarningMessage(
                `XLIDE: "${moduleName}" changed again after the agent's write, so reverting would ` +
                'discard the newer edit too. Use Review Agent Change to compare, then edit manually.',
            );
            return;
        }
        if (record.beforeExisted) {
            await deps.writeModuleSource(filePath, moduleName, record.before);
        } else {
            await deps.deleteModule(filePath, moduleName);
        }
        resolvePendingAgentReview(filePath, moduleName);
    } catch (err) {
        void vscode.window.showErrorMessage(
            `XLIDE: Could not revert "${moduleName}": ${errorMessage(err)}`,
        );
    }
}

/** A deleted module has nothing left to review. */
export function discardPendingAgentReview(filePath: string, moduleName: string): void {
    resolvePendingAgentReview(filePath, moduleName);
}

/** A renamed module carries its unreviewed agent change with it. */
export function renamePendingAgentReview(filePath: string, moduleName: string, newName: string): void {
    const record = pendingReviews.get(pendingKey(filePath, moduleName));
    if (!record) {
        return;
    }
    pendingReviews.delete(pendingKey(filePath, moduleName));
    pendingEmitter.fire({ filePath, moduleName });
    pendingReviews.set(pendingKey(filePath, newName), record);
    pendingEmitter.fire({ filePath, moduleName: newName });
}

/** The engine hides attribute headers from the editor surface; compare the
 * bodies the way the write path stores them - exact text, EOL-normalized. */
function normalizeForCompare(source: string): string {
    return source.replace(/\r\n?/g, '\n').trimEnd();
}

const BEFORE_IMAGE_CAP = 16;

function pruneBeforeImages(): void {
    while (beforeImages.size > BEFORE_IMAGE_CAP) {
        const oldest = beforeImages.keys().next().value;
        if (oldest === undefined) {
            return;
        }
        beforeImages.delete(oldest);
    }
}
