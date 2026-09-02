// Before-images for the last rename, so it can be put back as one operation.
//
// Issue #9 rule 10. A rename edits several modules at once, and an editor's
// undo stack is per document: undo in the file you are looking at reverses that
// file's share and leaves the rest renamed - a half-renamed project, from a
// keystroke that means "put it back". So the rename keeps the text it read out
// of each module immediately before writing, and offers an explicit Undo
// Rename that restores all of them together.
//
// One slot, deliberately. A stack would invite undoing a rename from three
// renames ago, by which time the before-images describe modules that have moved
// on and restoring them would discard whatever came after.
//
// Kept free of any `vscode` import so the bookkeeping can be unit-tested.

export interface RenamedModuleImage {
    moduleName: string;
    /** Module source exactly as read immediately before the rename was written. */
    before: string;
}

export interface RenameSnapshot {
    projectPath: string;
    oldName: string;
    newName: string;
    modules: RenamedModuleImage[];
    /** Set when the rename also renamed the component, not only its references. */
    renamedModule?: { from: string; to: string };
}

let lastRename: RenameSnapshot | undefined;
/**
 * Writes the rename itself is about to cause. Applying its edits goes through
 * the same file-system path as a developer pressing Save, so without this the
 * rename would immediately invalidate its own history and never offer an undo.
 * Each expected write is consumed once.
 */
let expectedWrites = new Set<string>();

function writeKey(projectPath: string, moduleName: string): string {
    return `${projectPath.toLowerCase()}::${moduleName.toLowerCase()}`;
}

/** Records the before-images of a rename that is about to be written. */
export function recordRename(snapshot: RenameSnapshot): void {
    lastRename = snapshot.modules.length > 0 || snapshot.renamedModule ? snapshot : undefined;
    expectedWrites = new Set(
        lastRename
            ? lastRename.modules.map((image) => writeKey(snapshot.projectPath, image.moduleName))
            : [],
    );
}

/**
 * Reports a module write. A write the rename was about to make is consumed and
 * changes nothing; any other write means the before-images no longer describe
 * what is on disk, so the history is dropped rather than left to overwrite
 * somebody else's change.
 */
export function noteModuleWrite(projectPath: string, moduleName: string): void {
    const key = writeKey(projectPath, moduleName);
    if (expectedWrites.delete(key)) {
        return;
    }
    invalidateRenameHistory(projectPath);
}

/** The rename that can currently be undone, if any. */
export function pendingUndo(): RenameSnapshot | undefined {
    return lastRename;
}

/**
 * Takes the snapshot for undoing, clearing it so the same rename cannot be
 * put back twice - the second attempt would write stale text over whatever the
 * first restore produced.
 */
export function takeRenameForUndo(): RenameSnapshot | undefined {
    const snapshot = lastRename;
    lastRename = undefined;
    expectedWrites = new Set();
    return snapshot;
}

/**
 * Drops the recorded rename. Called when something else writes to the project,
 * because the before-images no longer describe what is on disk and restoring
 * them would discard that other change.
 */
export function invalidateRenameHistory(projectPath?: string): void {
    if (!projectPath || !lastRename) {
        lastRename = undefined;
        expectedWrites = new Set();
        return;
    }
    if (lastRename.projectPath.toLowerCase() === projectPath.toLowerCase()) {
        lastRename = undefined;
        expectedWrites = new Set();
    }
}

/** Human-readable description of what Undo Rename would put back. */
export function describeUndo(snapshot: RenameSnapshot): string {
    const moduleCount = snapshot.modules.length;
    const where = moduleCount === 1 ? '1 module' : `${moduleCount} modules`;
    const component = snapshot.renamedModule
        ? `, and rename the module back to ${snapshot.renamedModule.from}`
        : '';
    return `Undo renaming '${snapshot.oldName}' to '${snapshot.newName}' across ${where}${component}.`;
}
