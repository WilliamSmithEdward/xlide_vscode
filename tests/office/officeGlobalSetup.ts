import * as fs from 'fs';
import * as path from 'path';
import { SCRATCH } from './officeHarness';

/** Starts each run with an empty scratch folder. */
export default function setup(): void {
    try {
        fs.rmSync(SCRATCH, { recursive: true, force: true });
    } catch {
        // Something still holds a file from the last run; the checks each
        // start from a fresh subfolder anyway.
    }
    fs.mkdirSync(SCRATCH, { recursive: true });
    // rmSync above can fail on a held fixture. Installation probes still
    // must never survive into a new run.
    for (const host of ['excel', 'word', 'powerpoint', 'access']) {
        fs.rmSync(path.join(SCRATCH, `installed-${host}.txt`), { force: true });
    }
}
