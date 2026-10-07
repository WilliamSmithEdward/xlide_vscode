import * as vscode from 'vscode';
import * as path from 'path';
import { projectIdentityKey } from './projectIdentity';
import { enginePriming } from './enginePriming';
import { hostPlatform } from './vba/hostPlatform';
import { assertProjectAccess, isProjectAccessLocked, unlockProject } from './vba/projectService';

// Explorer, editors and agents can arrive together; one user prompt per file.
const pending = new Map<string, Promise<void>>();
const unlocked = new vscode.EventEmitter<string>();
/** Raised only after successful authorization; listeners receive no password. */
export const onDidUnlockVbaProject = unlocked.event;

export async function ensureProjectPassword(filePath: string, allowMissing = false): Promise<void> {
    await enginePriming.prime([filePath]);
    if (allowMissing && !hostPlatform().exists(filePath)) { return; }
    if (!isProjectAccessLocked(filePath)) { return; }
    const key = projectIdentityKey(filePath);
    const existing = pending.get(key);
    if (existing) { return existing; }
    const request = promptForPassword(filePath);
    pending.set(key, request);
    try { await request; }
    finally { pending.delete(key); }
    // Recheck current protection after an async prompt/shared wait.
    assertProjectAccess(filePath);
}

async function promptForPassword(filePath: string): Promise<void> {
    let incorrect = false;
    while (isProjectAccessLocked(filePath)) {
        let password = await vscode.window.showInputBox({
            title: 'Unlock VBA project',
            prompt: `${incorrect ? 'Incorrect password. ' : ''}Enter the VBA project password for ${path.basename(filePath)}. Access lasts for this XLIDE session.`,
            password: true,
            ignoreFocusOut: true,
        });
        if (password === undefined) { throw new vscode.CancellationError(); }
        try {
            await enginePriming.prime([filePath]);
            if (unlockProject(filePath, password)) { unlocked.fire(filePath); return; }
        } finally { password = undefined; }
        incorrect = true;
    }
}
