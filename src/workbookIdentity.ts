import * as path from 'path';

export function workbookIdentityKey(
    workbookPath: string,
    platform: NodeJS.Platform = process.platform,
): string {
    return platform === 'win32'
        ? path.win32.normalize(workbookPath).toLowerCase()
        : path.posix.normalize(workbookPath);
}

export function sameWorkbookPath(
    a: string,
    b: string,
    platform: NodeJS.Platform = process.platform,
): boolean {
    return workbookIdentityKey(a, platform) === workbookIdentityKey(b, platform);
}

export function moduleIdentityKey(moduleName: string): string {
    return moduleName.toLowerCase();
}
