import { osPlatform, type OsPlatform } from './util/osPlatform';
import * as path from 'path';

export function projectIdentityKey(
    projectPath: string,
    platform: OsPlatform = osPlatform(),
): string {
    if (platform === 'win32') {
        return path.win32.normalize(projectPath).toLowerCase();
    }
    if (platform === 'web') {
        // The same file reaches this two ways in a browser, and they do not
        // look alike: uri.fsPath renders a virtual workspace's URI with
        // backslashes ('\Book.xlsm'), while decoding an xlide-vba:// module
        // URI keeps the URI's own forward slashes ('/Book.xlsm'). Both name
        // one workbook, so the key cannot depend on which form the caller
        // happened to be holding - that mismatch left the tree unable to
        // recognize its own active project, and it collapsed the workbook.
        //
        // Only the web branch does this. On a real POSIX filesystem a
        // backslash is an ordinary character in a filename, and rewriting it
        // would merge two genuinely different files.
        return path.posix.normalize(projectPath.replace(/\\/g, '/'));
    }
    return path.posix.normalize(projectPath);
}

export function sameProjectPath(
    a: string,
    b: string,
    platform: OsPlatform = osPlatform(),
): boolean {
    return projectIdentityKey(a, platform) === projectIdentityKey(b, platform);
}

export function moduleIdentityKey(moduleName: string): string {
    return moduleName.toLowerCase();
}
