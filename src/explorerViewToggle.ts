import { errorMessage } from './util/errors';
import type { XlideExplorerView } from './globalSettings';

export interface ExplorerViewToggleDeps {
    /** Switches the tree and publishes the context key the buttons read. */
    applyView(view: XlideExplorerView): void;
    /** Remembers the view in the user's settings. */
    persist(view: XlideExplorerView): Promise<unknown>;
    log(message: string): void;
    /** Told once per window that the view switched but could not be saved. */
    warnPersistFailed(): void;
}

/**
 * The Tree / Folders toggle.
 *
 * The view is switched first and remembered second. Writing the setting can
 * fail in a window whose configuration registry no longer holds the key, which
 * is what an extension update in a live window leaves behind: VS Code then
 * refuses the write, calling `xlide.explorer.view` an unregistered
 * configuration (issue #71). Switching first means the button still does its
 * job there, and only the remembering is lost until the window is reloaded.
 * The warning is shown once, because the next press would fail the same way.
 */
export function createExplorerViewSetter(
    deps: ExplorerViewToggleDeps,
): (view: XlideExplorerView) => Promise<void> {
    let warned = false;
    return async (view: XlideExplorerView): Promise<void> => {
        deps.applyView(view);
        try {
            await deps.persist(view);
        } catch (err) {
            deps.log(`Explorer view set to ${view} for this window only: ${errorMessage(err)}`);
            if (!warned) {
                warned = true;
                deps.warnPersistFailed();
            }
        }
    };
}
