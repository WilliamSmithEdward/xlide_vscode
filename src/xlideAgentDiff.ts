// Review surface for AI-agent module writes.
//
// When Copilot (or any agent) writes VBA through xlide_writeModule, the edit
// never passes through the editor's edit pipeline, so Copilot's own
// keep/reject review cannot see it and no diff appears anywhere. This module
// supplies the same contract from XLIDE's side: the write's before-image is
// kept, a native diff opens (before <-> the live module document), and a
// notification offers Keep / Revert - Revert restores the before-image
// through the normal audited write path, refusing when the module has
// drifted past the agent's write in the meantime.

import * as vscode from 'vscode';
import { encodeModuleUri, moduleIdentityKey, workbookIdentityKey } from './xlideFileSystem';
import { errorMessage } from './util/errors';

export const XLIDE_AGENT_BEFORE_SCHEME = 'xlide-vba-before';

interface AgentWriteRecord {
    before: string;
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
    return `${workbookIdentityKey(filePath)}::${moduleIdentityKey(moduleName)}`;
}

export function hasPendingAgentReview(filePath: string, moduleName: string): boolean {
    return pendingReviews.has(pendingKey(filePath, moduleName));
}

function setPendingAgentReview(filePath: string, moduleName: string, record: AgentWriteRecord): void {
    pendingReviews.set(pendingKey(filePath, moduleName), record);
    pendingEmitter.fire({ filePath, moduleName });
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
}

/**
 * Opens the before/current diff for an agent write and offers Keep / Revert.
 * Fire-and-forget: failures surface as messages, never into the tool result.
 */
export async function reviewAgentModuleWrite(
    deps: AgentWriteReviewDeps,
    filePath: string,
    moduleName: string,
    record: AgentWriteRecord,
): Promise<void> {
    setPendingAgentReview(filePath, moduleName, record);
    await runAgentReviewPrompt(deps, filePath, moduleName, record);
}

/**
 * Reopens the review for a module whose agent write is still pending -
 * the tree's "Review Agent Change" action. No-op when nothing is pending.
 */
export async function reopenAgentReview(
    deps: AgentWriteReviewDeps,
    filePath: string,
    moduleName: string,
): Promise<void> {
    const record = pendingReviews.get(pendingKey(filePath, moduleName));
    if (!record) {
        void vscode.window.showInformationMessage(
            `XLIDE: "${moduleName}" has no unreviewed agent change.`,
        );
        return;
    }
    await runAgentReviewPrompt(deps, filePath, moduleName, record);
}

async function runAgentReviewPrompt(
    deps: AgentWriteReviewDeps,
    filePath: string,
    moduleName: string,
    record: AgentWriteRecord,
): Promise<void> {
    const liveUri = encodeModuleUri(filePath, moduleName);
    // A fresh URI per write: an older diff view keeps showing its own frozen
    // before-image instead of silently changing under the reader.
    writeCounter += 1;
    const beforeUri = vscode.Uri.from({
        scheme: XLIDE_AGENT_BEFORE_SCHEME,
        path: `${liveUri.path}.${writeCounter}`,
    });
    beforeImages.set(beforeUri.toString(), record.before);
    pruneBeforeImages();

    try {
        await vscode.commands.executeCommand(
            'vscode.diff',
            beforeUri,
            liveUri,
            `${moduleName}: before agent edit ↔ current`,
            { preview: true },
        );
    } catch {
        // The diff view is best-effort; the Keep/Revert prompt still runs.
    }

    const choice = await vscode.window.showInformationMessage(
        `An AI agent wrote VBA module "${moduleName}". Keep the change?`,
        'Keep',
        'Revert',
    );
    if (choice === 'Keep') {
        resolvePendingAgentReview(filePath, moduleName);
        return;
    }
    if (choice !== 'Revert') {
        // Dismissed: the change stays pending, and the tree badge keeps the
        // review reachable from the module's context menu.
        return;
    }
    try {
        const current = await deps.readModuleSource(filePath, moduleName);
        if (normalizeForCompare(current) !== normalizeForCompare(record.after)) {
            void vscode.window.showWarningMessage(
                `XLIDE: "${moduleName}" changed again after the agent's write, so reverting would ` +
                'discard the newer edit too. Review the open diff and edit manually instead.',
            );
            return;
        }
        await deps.writeModuleSource(filePath, moduleName, record.before);
        resolvePendingAgentReview(filePath, moduleName);
        void vscode.window.showInformationMessage(`XLIDE: Reverted the agent's change to "${moduleName}".`);
    } catch (err) {
        void vscode.window.showErrorMessage(
            `XLIDE: Could not revert "${moduleName}": ${errorMessage(err)}`,
        );
    }
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
