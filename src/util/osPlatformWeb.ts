// The platform name in a browser.
//
// Every caller asks in order to decide between Windows rules and everything
// else: case-insensitive paths, drive letters, a local Office to coordinate
// with. A web workspace has none of those, so 'web' lands on the correct
// branch in each case while naming what it actually is.

import type { OsPlatform } from './osPlatform';

export const osPlatform: OsPlatform = 'web';
