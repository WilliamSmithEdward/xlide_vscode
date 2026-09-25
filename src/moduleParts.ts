// Parts of a module, for the agent tools: the lines a range or a procedure
// covers, several of them read in one call, and edits to several parts
// applied in one write (issue #95). An agent working on a large module read
// it whole, or one window at a time, and wrote it whole; here it names the
// parts it wants and gets, or changes, exactly those.
//
// Lines are 1-based and inclusive, numbered as xlide_readModule shows them.
// A procedure's lines are the ones the VBE's ProcOfLine gives it (the comment
// lines directly above its header, the header, through its End line), without
// the blank lines that separate it from the next procedure. No vscode
// dependency: the engine's source in, text out.

import { vbaProcedureRanges, type VbaProcedureRange } from './vbaProcedureAtLine';

/** A range of lines, 1-based and inclusive at both ends. */
export interface ModuleLineRange {
    startLine: number;
    endLine: number;
}

export interface ModulePartsRequest {
    ranges?: readonly ModuleLineRange[];
    /** Procedure names, or `Kind Name` where a name is shared, such as `Property Let Value`. */
    procedures?: readonly string[];
}

export interface ModulePart {
    /** `lines 10-20`, or `Sub Foo` for a procedure. */
    label: string;
    startLine: number;
    endLine: number;
    text: string;
}

export type ModuleEdit =
    | { startLine: number; endLine: number; text: string }
    | { insertAfterLine: number; text: string }
    | { procedure: string; text: string };

export interface AppliedModuleEdit {
    label: string;
    /** Where the edit landed in the module as it now reads; undefined for a deletion. */
    newStartLine?: number;
    newEndLine?: number;
}

export type ModulePartsResult =
    | { ok: true; parts: ModulePart[] }
    | { ok: false; message: string };

/** The line above a part in a tool result: what it is, and which lines it is. */
export function modulePartHeading(part: ModulePart): string {
    return part.label.startsWith('lines ')
        ? `--- ${part.label}`
        : `--- ${part.label} (lines ${part.startLine}-${part.endLine})`;
}

export type ModuleEditsResult =
    | { ok: true; source: string; applied: AppliedModuleEdit[] }
    | { ok: false; message: string };

/** CRLF, lone CR and lone LF: the splits the procedure ranges use too. */
const LINE_BREAK_RE = /\r\n|\r|\n/;

/**
 * The module's lines as the tools number them. The engine's text ends with a
 * line break, and what follows it is not a line.
 */
export function splitModuleLines(source: string): string[] {
    const lines = source.split(LINE_BREAK_RE);
    if (lines.length > 1 && lines[lines.length - 1] === '') {
        lines.pop();
    }
    return lines;
}

/** The line break the module uses, for the text that goes back to it. */
function lineBreakOf(source: string): string {
    return source.includes('\r\n') ? '\r\n' : '\n';
}

/** The lines of an edit's text. Empty text is no lines: a range given it is deleted. */
function textLines(text: string): string[] {
    return text === '' ? [] : splitModuleLines(text);
}

/** `Name`, or `Kind Name`, where the kind of a property can be given as `Get`, `Let` or `Set` alone. */
const PROCEDURE_SPEC_RE =
    /^\s*(?:(Sub|Function|Property\s+(?:Get|Let|Set)|Get|Let|Set)\s+)?([\p{L}_][\p{L}\p{M}\p{N}_]*)\s*$/iu;

/** The procedures of the module, with the lines each one covers (0-based, inclusive). */
function procedureSpans(lines: readonly string[], source: string): VbaProcedureRange[] {
    return vbaProcedureRanges(source).map((range) => {
        // ProcOfLine gives a procedure the blank lines above its comments and
        // down to the next procedure, and the last procedure everything to
        // the end of the module. The blank lines separate; they are nobody's.
        let first = range.firstLine;
        while (first < range.lastLine && lines[first].trim() === '') {
            first += 1;
        }
        let last = Math.min(range.lastLine, lines.length - 1);
        while (last > first && lines[last].trim() === '') {
            last -= 1;
        }
        return { ...range, firstLine: first, lastLine: last };
    });
}

function procedureLabel(range: VbaProcedureRange): string {
    return `${range.kind} ${range.name}`;
}

/** The procedure a spec names, or the reason it names none or several. */
function findProcedure(
    spans: readonly VbaProcedureRange[],
    spec: string,
): { range: VbaProcedureRange } | { message: string } {
    const match = PROCEDURE_SPEC_RE.exec(spec);
    if (!match) {
        return { message: `"${spec}" is not a procedure name. Give the name, or the kind and name, such as "Property Let Value".` };
    }
    const given = match[1]?.replace(/\s+/g, ' ').toLowerCase();
    const kind = given !== undefined && /^(get|let|set)$/.test(given) ? `property ${given}` : given;
    const name = match[2].toLowerCase();
    const found = spans.filter((range) =>
        range.name.toLowerCase() === name && (kind === undefined || range.kind.toLowerCase() === kind));
    if (found.length === 1) {
        return { range: found[0] };
    }
    if (found.length === 0) {
        return { message: `The module has no procedure "${spec}". xlide_listSubs lists its procedures.` };
    }
    return {
        message: `"${spec}" names ${found.length} procedures: ${found.map(procedureLabel).join(', ')}. Give the kind too.`,
    };
}

function checkRange(range: ModuleLineRange, lineCount: number): string | undefined {
    if (!Number.isInteger(range.startLine) || !Number.isInteger(range.endLine)) {
        return 'startLine and endLine must be whole numbers.';
    }
    if (range.startLine < 1 || range.endLine < range.startLine) {
        return `lines ${range.startLine}-${range.endLine} is not a range: startLine is 1 or more, and endLine is not before it.`;
    }
    if (range.endLine > lineCount) {
        return `lines ${range.startLine}-${range.endLine} run past the end of the module, which has ${lineCount} lines.`;
    }
    return undefined;
}

/**
 * The parts of a module a read names, in the order they were asked for.
 * A range that runs past the module and a procedure the module lacks are
 * refused whole: a part silently cut short would be read as the real thing.
 */
export function readModuleParts(source: string, request: ModulePartsRequest): ModulePartsResult {
    const lines = splitModuleLines(source);
    const parts: ModulePart[] = [];
    for (const range of request.ranges ?? []) {
        const problem = checkRange(range, lines.length);
        if (problem) {
            return { ok: false, message: problem };
        }
        parts.push({
            label: `lines ${range.startLine}-${range.endLine}`,
            startLine: range.startLine,
            endLine: range.endLine,
            text: lines.slice(range.startLine - 1, range.endLine).join('\n'),
        });
    }
    const spans = request.procedures?.length ? procedureSpans(lines, source) : [];
    for (const spec of request.procedures ?? []) {
        const found = findProcedure(spans, spec);
        if ('message' in found) {
            return { ok: false, message: found.message };
        }
        parts.push({
            label: procedureLabel(found.range),
            startLine: found.range.firstLine + 1,
            endLine: found.range.lastLine + 1,
            text: lines.slice(found.range.firstLine, found.range.lastLine + 1).join('\n'),
        });
    }
    return { ok: true, parts };
}

/** An edit resolved against the module as it is: the lines it takes out, and what it puts in. */
interface ResolvedEdit {
    label: string;
    /** 0-based, inclusive start. */
    start: number;
    /** 0-based, exclusive end; equal to start for an insertion. */
    end: number;
    lines: string[];
}

function resolveEdit(
    edit: ModuleEdit,
    lines: readonly string[],
    spans: () => readonly VbaProcedureRange[],
): ResolvedEdit | { message: string } {
    if (!edit || typeof edit !== 'object' || typeof edit.text !== 'string') {
        return { message: 'an edit is an object with a text field, plus startLine and endLine, insertAfterLine, or procedure.' };
    }
    if ('procedure' in edit) {
        const found = findProcedure(spans(), edit.procedure);
        if ('message' in found) {
            return found;
        }
        return {
            label: procedureLabel(found.range),
            start: found.range.firstLine,
            end: found.range.lastLine + 1,
            lines: textLines(edit.text),
        };
    }
    if ('insertAfterLine' in edit) {
        const after = edit.insertAfterLine;
        if (!Number.isInteger(after) || after < 0 || after > lines.length) {
            return { message: `insertAfterLine ${after} is not a line of the module: 0 inserts at the top, and ${lines.length} at the end.` };
        }
        return { label: `after line ${after}`, start: after, end: after, lines: textLines(edit.text) };
    }
    const problem = checkRange(edit, lines.length);
    if (problem) {
        return { message: problem };
    }
    return {
        label: `lines ${edit.startLine}-${edit.endLine}`,
        start: edit.startLine - 1,
        end: edit.endLine,
        lines: textLines(edit.text),
    };
}

/**
 * Applies edits to a module in one pass. Every line number means the module
 * as it was read, whatever the edits above it do to the numbering, so the
 * edits are applied from the bottom up. Two edits that touch the same lines,
 * or two insertions at the same place, are refused: which would come first
 * is a guess.
 */
export function applyModuleEdits(source: string, edits: readonly ModuleEdit[]): ModuleEditsResult {
    if (edits.length === 0) {
        return { ok: false, message: 'edits is empty: nothing to change.' };
    }
    const lines = splitModuleLines(source);
    let cachedSpans: readonly VbaProcedureRange[] | undefined;
    const spans = (): readonly VbaProcedureRange[] => (cachedSpans ??= procedureSpans(lines, source));
    const resolved: Array<ResolvedEdit & { index: number }> = [];
    for (const [index, edit] of edits.entries()) {
        const one = resolveEdit(edit, lines, spans);
        if ('message' in one) {
            return { ok: false, message: `edits[${index}]: ${one.message}` };
        }
        resolved.push({ ...one, index });
    }
    const ordered = [...resolved].sort((a, b) => a.start - b.start || a.end - b.end);
    for (let i = 1; i < ordered.length; i += 1) {
        const previous = ordered[i - 1];
        const current = ordered[i];
        const overlap = current.start < previous.end
            || (current.start === previous.start && current.end === previous.end);
        if (overlap) {
            return {
                ok: false,
                message: `edits[${previous.index}] (${previous.label}) and edits[${current.index}] (${current.label}) touch the same lines. Make them one edit.`,
            };
        }
    }
    const out = [...lines];
    for (const edit of [...ordered].reverse()) {
        out.splice(edit.start, edit.end - edit.start, ...edit.lines);
    }
    // Where each edit landed, walking down with the lines the edits above
    // it added or removed.
    let shift = 0;
    const landed = new Map<number, AppliedModuleEdit>();
    for (const edit of ordered) {
        const newStart = edit.start + shift;
        landed.set(edit.index, edit.lines.length === 0
            ? { label: edit.label }
            : { label: edit.label, newStartLine: newStart + 1, newEndLine: newStart + edit.lines.length });
        shift += edit.lines.length - (edit.end - edit.start);
    }
    const lineBreak = lineBreakOf(source);
    const joined = out.join(lineBreak);
    return {
        ok: true,
        source: joined === '' ? '' : `${joined}${lineBreak}`,
        applied: resolved.map((edit) => landed.get(edit.index)!),
    };
}
