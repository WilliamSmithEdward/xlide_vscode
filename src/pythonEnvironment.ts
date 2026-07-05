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
