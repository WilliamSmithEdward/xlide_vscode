// Failures the analyser recovered from, written to XLIDE's output channel
// (issue #178). Analysis never throws, so a module it could not fully check
// looked the same as a clean one: fewer problems, or none, with nothing said.
// Each failure is written once per module, with its stack, so a report can
// carry it; analysis runs again on every keystroke and would repeat it.

import type { VbaModuleAnalysisFailure } from './vbaModuleAnalysis';

const MAX_REMEMBERED = 500;

let sink: ((line: string) => void) | undefined;
const written = new Set<string>();

/** Where failures are written; the extension sets its output channel. */
export function setAnalysisFailureLog(log: ((line: string) => void) | undefined): void {
    sink = log;
}

/** What a failure cost, in words. */
export function describeAnalysisFailure(failure: VbaModuleAnalysisFailure): string {
    switch (failure.stage) {
        case 'options':
            return `an option was left out: ${failure.message}`;
        case 'analysis':
            return `no rule checked it: ${failure.message}`;
        case 'rule':
            return `rule ${failure.rule} did not run: ${failure.message}`;
        case 'statement-walk':
        case 'expression-walk':
            return failure.rule
                ? `rule ${failure.rule} stopped partway: ${failure.message}`
                : `the ${failure.stage === 'statement-walk' ? 'statement' : 'expression'} rules stopped partway: ${failure.message}`;
        case 'index':
            return `it was left out of its project, so no other module sees its declarations: ${failure.message}`;
        case 'project-context':
            return `it was checked without the rest of its project: ${failure.message}`;
        case 'test-directives':
            return `its test directives were not checked: ${failure.message}`;
        case 'structural':
            return `its block structure was not checked: ${failure.message}`;
    }
}

/** Writes each failure not yet written for this module. */
export function logAnalysisFailures(module: string, failures: readonly VbaModuleAnalysisFailure[] | undefined): void {
    if (!sink || !failures) {
        return;
    }
    for (const failure of failures) {
        const key = `${module}\n${failure.stage}\n${failure.rule ?? ''}\n${failure.message}`;
        if (written.has(key)) {
            continue;
        }
        if (written.size >= MAX_REMEMBERED) {
            written.clear();
        }
        written.add(key);
        sink(`XLIDE could not fully check ${module}: ${describeAnalysisFailure(failure)}`
            + (failure.stack ? `\n${failure.stack}` : ''));
    }
}
