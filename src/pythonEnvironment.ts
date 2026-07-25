// Pure helpers for classifying Python-backend startup failures and locating
// interpreters. Kept free of any `vscode` dependency so they can be unit-tested
// without the VS Code module graph.

/**
 * True when a backend-start error message means "no usable Python interpreter".
 *
 * Beyond the plain spawn failures (ENOENT, cmd.exe "not recognized"), this must
 * recognize the two OS-level Python STUBS that make the spawn itself succeed and
 * then fail with a distinctive message:
 * - Windows Store app-execution alias: "Python was not found; run without
 *   arguments to install from the Microsoft Store..."
 * - macOS Command Line Tools stub: /usr/bin/python3 exists on every macOS but,
 *   without the developer tools installed, dies with "xcrun: error: invalid
 *   active developer path (/Library/Developer/CommandLineTools)..." or an
 *   xcode-select "no developer tools were found" note.
 */
export function isPythonNotFoundMessage(msg: string): boolean {
    return /python.*not found|not recognized|cannot find|no such file|ENOENT|spawn.*python/i.test(msg)
        || isMacCltStubMessage(msg);
}

/** True when the failure is specifically the macOS Command Line Tools python3 stub. */
export function isMacCltStubMessage(msg: string): boolean {
    return /xcrun|xcode-select|invalid active developer path|developer tools were found|CommandLineTools/i.test(msg);
}

/** True when Python started but a required package (pyOpenVBA, openpyxl) is missing. */
export function isMissingPackageMessage(msg: string): boolean {
    return /No module named|ModuleNotFoundError|ImportError/i.test(msg);
}

/**
 * Well-known interpreter locations to prefer over a bare `python3` on POSIX
 * platforms. On macOS the bare name can resolve to the useless Command Line
 * Tools stub at /usr/bin/python3, and Homebrew's directory is often not on the
 * extension host's PATH at all - so a user who installs Python after seeing the
 * setup warning is only detected if these locations are probed directly.
 * Ordered: Homebrew (Apple Silicon), Homebrew (Intel) / python.org symlink,
 * python.org framework install.
 */
/** One tracked runtime library that has a newer release than the installed one. */
export interface OutdatedPythonLibrary {
    name: string;
    installed: string;
    latest: string;
}

/**
 * PEP-440-ish version comparison over dot-separated segments. Numeric segments
 * compare numerically; a shorter version is padded with zeros (3.1 == 3.1.0).
 * Any non-numeric segment (a pre-release like "3.2.0rc1") makes the comparison
 * conservative: the versions are treated as equal so an rc/dev build never
 * produces an "update available" nudge in either direction.
 */
export function compareVersionStrings(a: string, b: string): number {
    const segsA = a.trim().split('.');
    const segsB = b.trim().split('.');
    // Any non-numeric segment on either side (a pre-release/dev build) makes
    // the whole comparison indeterminate - checked up front, so a numeric
    // difference in an earlier segment cannot decide against a "3.2.0rc1".
    if ([...segsA, ...segsB].some((seg) => !/^\d+$/.test(seg))) {
        return 0;
    }
    const len = Math.max(segsA.length, segsB.length);
    for (let i = 0; i < len; i++) {
        const numA = Number(segsA[i] ?? '0');
        const numB = Number(segsB[i] ?? '0');
        if (numA !== numB) {
            return numA < numB ? -1 : 1;
        }
    }
    return 0;
}

/**
 * The tracked libraries whose installed version is strictly behind the latest
 * released one. Unknown entries on either side are skipped: no data, no nudge.
 */
export function outdatedPythonLibraries(
    installed: Readonly<Record<string, string>>,
    latest: Readonly<Record<string, string>>,
): OutdatedPythonLibrary[] {
    const out: OutdatedPythonLibrary[] = [];
    for (const [name, installedVersion] of Object.entries(installed)) {
        const latestVersion = latest[name];
        if (!installedVersion || !latestVersion) {
            continue;
        }
        if (compareVersionStrings(installedVersion, latestVersion) < 0) {
            out.push({ name, installed: installedVersion, latest: latestVersion });
        }
    }
    return out;
}

export function posixPythonCandidatePaths(platform: NodeJS.Platform): string[] {
    if (platform !== 'darwin') {
        // Linux: a real /usr/bin/python3 resolves via PATH, and a missing one
        // fails with ENOENT (already classified) - no candidates needed.
        return [];
    }
    return [
        '/opt/homebrew/bin/python3',
        '/usr/local/bin/python3',
        '/Library/Frameworks/Python.framework/Versions/Current/bin/python3',
    ];
}
