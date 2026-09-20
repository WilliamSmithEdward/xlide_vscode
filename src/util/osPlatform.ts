// Which platform XLIDE is running on, without reaching for `process`.
//
// A browser has no `process` at all, so `process.platform` throws rather than
// returning something unhelpful. The browser build reports 'web', which is
// the right answer everywhere this is asked: a web workspace's paths are
// POSIX and case-sensitive, and there is no local Office to coordinate with.

import { osPlatform as leaf } from './osPlatformNode';

/** Node's platform names, plus the browser. */
export type OsPlatform = NodeJS.Platform | 'web';

export const osPlatform: OsPlatform = leaf;

/** Windows path rules: case-insensitive, drive letters, backslashes. */
export const isWindows = osPlatform === 'win32';
