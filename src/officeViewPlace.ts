// Where the user is in a copy of the file that XLIDE closes and opens again.
//
// XLIDE reopens a copy of the file behind a save in two places: the read-only
// refresh after a module save (Excel, where a read-only copy does not lock
// the file, so the save goes through and the copy goes stale), and the reopen
// after a save had to close the copy to free the lock (Word and PowerPoint
// read-only copies, and any copy under closeTracked or closeForce). Either
// way the fresh copy opened at the top of the first sheet, page or slide, so
// a read-only window meant to follow your edits threw away your place on
// every save. The same few lines now capture it before the close and put it
// back after the open:
//
//   - Excel: the active sheet, the selection and its active cell, and the
//     window's scroll position.
//   - Word: the selection, and how far down the window is scrolled.
//   - PowerPoint: the slide in view.
//
// And in each, which of the application's other files was in front: opening
// a file makes it the active one, which it was not if you had moved on.
//
// Plain strings, like officeFileState.ts, so the browser build carries them
// harmlessly; only the desktop runs them. Every step is best-effort: a place
// that cannot be read or put back leaves the copy where the open put it.

import { psSingleQuoted } from './util/powershell';
import type { OfficeHostApp } from './officeHostApps';

/** A place in a copy of the file, as the capture lines read it. */
export interface OfficeViewPlace {
    /** Excel: the active sheet's name. */
    sheet?: string;
    /** Excel: the selected range's address. */
    selection?: string;
    /** Excel: the active cell's address, inside the selection. */
    activeCell?: string;
    /** Excel: the window's first visible row. */
    scrollRow?: number;
    /** Excel: the window's first visible column. */
    scrollColumn?: number;
    /** Word: the selection's start, in characters. */
    start?: number;
    /** Word: the selection's end, in characters. */
    end?: number;
    /** Word: how far down the window is scrolled, as a percentage. */
    scrolled?: number;
    /** PowerPoint: the index of the slide in view. */
    slide?: number;
    /** The application's other file that was in front, by full name. */
    otherActive?: string;
}

const PLACE_SENTINEL = 'XLIDE_PLACE|';

/** The collection each application keeps its open files in, and how one is brought forward. */
const OPEN_FILES: Record<Exclude<OfficeHostApp, 'access'>, { collection: string; active: string; activate: string }> = {
    excel: { collection: 'Workbooks', active: 'ActiveWorkbook', activate: '$other.Activate()' },
    word: { collection: 'Documents', active: 'ActiveDocument', activate: '$other.Activate()' },
    powerpoint: { collection: 'Presentations', active: 'ActivePresentation', activate: '$other.Windows.Item(1).Activate()' },
};

const CAPTURE: Record<Exclude<OfficeHostApp, 'access'>, readonly string[]> = {
    excel: [
        'try {',
        '  $window = $file.Windows.Item(1)',
        '  $place.sheet = [string]$window.ActiveSheet.Name',
        '  $place.scrollRow = [int]$window.ScrollRow',
        '  $place.scrollColumn = [int]$window.ScrollColumn',
        '  try { $place.selection = [string]$window.RangeSelection.Address() } catch { }',
        '  try { $place.activeCell = [string]$window.ActiveCell.Address() } catch { }',
        '} catch { }',
    ],
    word: [
        'try {',
        '  $window = $file.ActiveWindow',
        '  $place.start = [int]$window.Selection.Start',
        '  $place.end = [int]$window.Selection.End',
        '  $place.scrolled = [int]$window.VerticalPercentScrolled',
        '} catch { }',
    ],
    powerpoint: [
        'try { $place.slide = [int]$file.Windows.Item(1).View.Slide.SlideIndex } catch { }',
    ],
};

const RESTORE: Record<Exclude<OfficeHostApp, 'access'>, readonly string[]> = {
    excel: [
        // Sheet, then selection, then the active cell inside it, then the
        // scroll: selecting scrolls the window to the selection.
        'try { if ($place.ContainsKey("sheet")) { $file.Sheets.Item($place.sheet).Activate() } } catch { }',
        'try { if ($place.ContainsKey("selection")) { $file.ActiveSheet.Range($place.selection).Select() } } catch { }',
        'try { if ($place.ContainsKey("activeCell")) { $file.ActiveSheet.Range($place.activeCell).Activate() } } catch { }',
        'try {',
        '  $window = $file.Windows.Item(1)',
        '  if ($place.ContainsKey("scrollRow")) { $window.ScrollRow = $place.scrollRow }',
        '  if ($place.ContainsKey("scrollColumn")) { $window.ScrollColumn = $place.scrollColumn }',
        '} catch { }',
    ],
    word: [
        'try {',
        '  $window = $file.ActiveWindow',
        '  if ($place.ContainsKey("start")) { $window.Selection.SetRange($place.start, $place.end) }',
        '  if ($place.ContainsKey("scrolled")) { $window.VerticalPercentScrolled = $place.scrolled }',
        '} catch { }',
    ],
    powerpoint: [
        'try { if ($place.ContainsKey("slide")) { $file.Windows.Item(1).View.GotoSlide($place.slide) } } catch { }',
    ],
};

/**
 * PowerShell reading the place in `$file`, the copy about to be closed, into
 * a hashtable `$place`. `$app` is its application. Nothing for Access, which
 * has no read-only copy to follow and reopens a database as it was.
 */
export function capturePlaceLines(host: OfficeHostApp): string[] {
    if (host === 'access') {
        return [];
    }
    const files = OPEN_FILES[host];
    return [
        '$place = @{}',
        ...CAPTURE[host],
        `try { $active = $app.${files.active}; if ($active -and ($active.FullName -ine $file.FullName)) { $place.otherActive = [string]$active.FullName } } catch { }`,
    ];
}

/**
 * PowerShell putting the hashtable `$place` back into `$file`, the copy just
 * opened, and then bringing back the application's other file that was in
 * front, if there was one.
 */
export function restorePlaceLines(host: OfficeHostApp): string[] {
    if (host === 'access') {
        return [];
    }
    const files = OPEN_FILES[host];
    return [
        'if ($place) {',
        ...RESTORE[host].map((line) => `  ${line}`),
        '  try {',
        '    if ($place.ContainsKey("otherActive")) {',
        `      foreach ($other in @($app.${files.collection})) { if ($other.FullName -ieq $place.otherActive) { ${files.activate}; break } }`,
        '    }',
        '  } catch { }',
        '}',
    ];
}

/** PowerShell printing `$place` for the caller to hand to the reopen. */
export const PLACE_REPORT_LINE =
    `if ($place) { [Console]::Out.WriteLine("${PLACE_SENTINEL}" + (ConvertTo-Json -InputObject $place -Compress)) }`;

const NUMBER_KEYS = ['scrollRow', 'scrollColumn', 'start', 'end', 'scrolled', 'slide'] as const;
const TEXT_KEYS = ['sheet', 'selection', 'activeCell', 'otherActive'] as const;

/** The place a script printed with {@link PLACE_REPORT_LINE}, or undefined. */
export function parsePlace(stdoutLines: readonly string[]): OfficeViewPlace | undefined {
    const line = stdoutLines.find((candidate) => candidate.startsWith(PLACE_SENTINEL));
    if (!line) {
        return undefined;
    }
    let raw: unknown;
    try {
        raw = JSON.parse(line.slice(PLACE_SENTINEL.length));
    } catch {
        return undefined;
    }
    if (!raw || typeof raw !== 'object') {
        return undefined;
    }
    const source = raw as Record<string, unknown>;
    const place: OfficeViewPlace = {};
    for (const key of NUMBER_KEYS) {
        const value = source[key];
        if (typeof value === 'number' && Number.isFinite(value)) {
            place[key] = Math.trunc(value);
        }
    }
    for (const key of TEXT_KEYS) {
        const value = source[key];
        if (typeof value === 'string' && value.length > 0) {
            place[key] = value;
        }
    }
    return Object.keys(place).length > 0 ? place : undefined;
}

/** PowerShell setting `$place` to a place read earlier, for {@link restorePlaceLines}. */
export function placeAssignmentLine(place: OfficeViewPlace): string {
    const entries: string[] = [];
    for (const key of NUMBER_KEYS) {
        const value = place[key];
        if (typeof value === 'number' && Number.isFinite(value)) {
            entries.push(`${key} = ${Math.trunc(value)}`);
        }
    }
    for (const key of TEXT_KEYS) {
        const value = place[key];
        if (typeof value === 'string' && value.length > 0) {
            entries.push(`${key} = ${psSingleQuoted(value)}`);
        }
    }
    return `$place = @{ ${entries.join('; ')} }`;
}
