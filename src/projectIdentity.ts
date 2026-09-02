import * as path from 'path';

export function projectIdentityKey(
    projectPath: string,
    platform: NodeJS.Platform = process.platform,
): string {
    return platform === 'win32'
        ? path.win32.normalize(projectPath).toLowerCase()
        : path.posix.normalize(projectPath);
}

export function sameProjectPath(
    a: string,
    b: string,
    platform: NodeJS.Platform = process.platform,
): boolean {
    return projectIdentityKey(a, platform) === projectIdentityKey(b, platform);
}

export function moduleIdentityKey(moduleName: string): string {
    return moduleName.toLowerCase();
}
