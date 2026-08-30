import { readFileSync } from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';

// F5 on a form writes the launcher module and then runs it. Both halves must
// sit INSIDE withWorkbookReopenSuppressed, exactly as the Run-Macro command
// does: otherwise the write's own post-save reopen (or its close-mode reopen)
// opens the workbook while the macro host opens it too, and one keypress
// produces two Excel windows.
//
// This reads the source because the wiring is what regressed - the helper
// itself is covered by excelWorkbookCoordinator.test.ts, and the command
// cannot run outside a real extension host.
const LAUNCH_SOURCE = readFileSync(path.join(__dirname, '..', 'src', 'vbaFormPreview.ts'), 'utf8');

describe('the F5 form launch', () => {
	it('suppresses the reopen across BOTH the write and the macro run', () => {
		const suppress = LAUNCH_SOURCE.indexOf('withWorkbookReopenSuppressed(');
		const write = LAUNCH_SOURCE.indexOf('runWriteWithExcelCoordination(');
		const run = LAUNCH_SOURCE.indexOf('runWorkbookMacroReadOnly(');
		expect(suppress).toBeGreaterThan(-1);
		expect(write).toBeGreaterThan(suppress);
		expect(run).toBeGreaterThan(suppress);
	});

	it('runs one launch at a time, so a second F5 cannot stack another', () => {
		expect(LAUNCH_SOURCE).toContain('if (launchInFlight) { return; }');
		expect(LAUNCH_SOURCE).toContain('launchInFlight = false;');
	});

	it('tracks the workbook the macro host reopened, so a later save can free it', () => {
		expect(LAUNCH_SOURCE).toContain('markWorkbookOpenedByXlide(wbPath)');
	});
});
