// One way to land bytes on disk: on a real filesystem a temp file beside the
// target and a rename over it, so a crash mid-write cannot leave a
// half-written file. Every project save and every VB6 module write goes
// through here.
//
// The mechanism is the platform's, because there is no rename in a browser -
// see hostPlatform.ts. This stays as the name the engine calls.

import { hostPlatform } from './hostPlatform';

export function atomicWrite(filePath: string, data: Buffer): void {
	hostPlatform().writeFile(filePath, data);
}
