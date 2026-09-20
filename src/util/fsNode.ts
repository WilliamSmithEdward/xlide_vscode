// Desktop existence check. The browser build never reaches this module -
// webBuild.js aliases it to fsWeb.ts - which is what keeps node:fs out of the
// web bundle while tests keep running against the real filesystem.

import * as fs from 'fs';

export async function fileExists(filePath: string): Promise<boolean> {
    try {
        await fs.promises.access(filePath, fs.constants.F_OK);
        return true;
    } catch {
        return false;
    }
}
