/**
 * The folder layout: the arithmetic that turns a flat module list plus each
 * module's `@Folder` annotation into the nested shape the explorer draws.
 *
 * Pure on purpose - the explorer supplies module nodes, this decides only
 * where they sit. Parity with xlide_vbide's foldertree.
 */

export interface FolderTreeModule {
    /** The normalized dotted path; absent or empty puts the module at the root. */
    folder?: string;
}

export interface FolderTreeFolder<T> {
    /** This segment's own name, in the first spelling seen. */
    name: string;
    /** The whole dotted path down to here - the annotation that makes it. */
    path: string;
    folders: FolderTreeFolder<T>[];
    modules: T[];
    /** Modules in this folder and in every folder under it. */
    moduleCount: number;
}

export interface FolderTree<T> {
    folders: FolderTreeFolder<T>[];
    /** Modules with no annotation, which sit at the project's own root. */
    modules: T[];
}

interface MutableFolder<T> {
    name: string;
    path: string;
    children: Map<string, MutableFolder<T>>;
    modules: T[];
}

/**
 * Modules keep the order they arrive in, which is the flat tree's order.
 * Folders sort by name without regard to case and come first, at every level.
 * Two annotations differing only in case are one folder, spelled the way it
 * was first seen.
 */
export function buildFolderTree<T extends FolderTreeModule>(modules: readonly T[]): FolderTree<T> {
    const roots = new Map<string, MutableFolder<T>>();
    const rootModules: T[] = [];

    for (const module of modules) {
        const segments = (module.folder ?? '').split('.').filter((segment) => segment.length > 0);
        if (segments.length === 0) {
            rootModules.push(module);
            continue;
        }
        let level = roots;
        let folder: MutableFolder<T> | undefined;
        let path = '';
        for (const segment of segments) {
            const key = segment.toLowerCase();
            path = path ? `${path}.${segment}` : segment;
            let next = level.get(key);
            if (!next) {
                // First spelling seen wins, so the path is built from the
                // folder's own name rather than from this module's annotation.
                next = { name: segment, path, children: new Map(), modules: [] };
                level.set(key, next);
            }
            path = next.path;
            folder = next;
            level = next.children;
        }
        folder!.modules.push(module);
    }

    return { folders: sortedFolders(roots), modules: rootModules };
}

function sortedFolders<T>(level: Map<string, MutableFolder<T>>): FolderTreeFolder<T>[] {
    return [...level.values()]
        .map((folder): FolderTreeFolder<T> => {
            const folders = sortedFolders(folder.children);
            return {
                name: folder.name,
                path: folder.path,
                folders,
                modules: folder.modules,
                moduleCount: folder.modules.length
                    + folders.reduce((total, child) => total + child.moduleCount, 0),
            };
        })
        .sort(compareFolderNames);
}

function compareFolderNames<T>(left: FolderTreeFolder<T>, right: FolderTreeFolder<T>): number {
    // 'accent' ignores case but keeps accented names apart, so "Ledger" and
    // "ledger" tie and the raw name breaks the tie deterministically.
    const byName = left.name.localeCompare(right.name, undefined, { sensitivity: 'accent' });
    return byName !== 0 ? byName : left.name.localeCompare(right.name);
}

/**
 * Every folder on the way to a module's own folder, outermost first:
 * "Accounts.Ledger" gives ["Accounts", "Accounts.Ledger"]. This is the chain
 * the tree opens when the editor moves to that module.
 */
export function folderPathChain(folder: string | undefined): string[] {
    const segments = (folder ?? '').split('.').filter((segment) => segment.length > 0);
    const chain: string[] = [];
    let path = '';
    for (const segment of segments) {
        path = path ? `${path}.${segment}` : segment;
        chain.push(path);
    }
    return chain;
}
