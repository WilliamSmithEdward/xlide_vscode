// The desktop platform name. The browser build never reaches this module -
// webBuild.js aliases it to osPlatformWeb.ts - because a browser has no
// `process` to ask.

import type { OsPlatform } from './osPlatform';

export const osPlatform: OsPlatform = process.platform;
