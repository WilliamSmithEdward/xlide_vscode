// The desktop platform name. The browser build never reaches this module -
// webBuild.js aliases it to osPlatformWeb.ts - because a browser has no
// `process` to ask.
//
// Read live on every call rather than captured once: see the note in
// osPlatform.ts about the coordination tests redefining process.platform.

import type { OsPlatform } from './osPlatform';

export function osPlatform(): OsPlatform {
    return process.platform;
}
