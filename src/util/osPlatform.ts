// Which platform XLIDE is running on, without reaching for `process`.
//
// A browser has no `process` at all, so `process.platform` throws rather than
// returning something unhelpful. The browser build reports 'web', which is
// the right answer everywhere this is asked: a web workspace's paths are
// POSIX and case-sensitive, and there is no local Office to coordinate with.
//
// Deliberately a function, not a constant. `process.platform` never changes
// at run time, so a constant would be correct in production - but the tests
// for the Windows-only Office coordination redefine `process.platform` to
// exercise that code on any machine, and a constant captured at import time
// freezes before they can. That is not a theoretical concern: making this a
// constant turned those tests green on Windows and red on CI's Linux runner.

import { osPlatform as leaf } from './osPlatformNode';

/** Node's platform names, plus the browser. */
export type OsPlatform = NodeJS.Platform | 'web';

export function osPlatform(): OsPlatform {
    return leaf();
}

/** Windows path rules: case-insensitive, drive letters, backslashes. */
export function isWindows(): boolean {
    return osPlatform() === 'win32';
}
