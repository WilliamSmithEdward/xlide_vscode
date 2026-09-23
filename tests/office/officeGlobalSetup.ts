import * as fs from 'fs';
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
}
