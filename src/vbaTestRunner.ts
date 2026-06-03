import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import { parseModule } from './analyzer/parser/parseModule';
import type { ProcedureNode } from './analyzer/parser/nodes';
import { lineStartOffsets } from './vbaStructuralAnalysis';
import { compareVbaModulesForTreeOrder } from './moduleDisplay';

export const XLIDE_VBA_TEST_DIRECTIVE = '@xlide-test';

export type VbaTestStatus = 'passed' | 'failed' | 'skipped';

export interface VbaTestModuleEntry {
    name: string;
    type: string;
    source?: string;
}

export interface VbaTestCase {
    id: string;
    moduleName: string;
    moduleType: string;
    procedureName: string;
    qualifiedName: string;
    line: number;
    column: number;
    annotationLine: number;
}

export interface VbaTestDiscoveryResult {
    filePath: string;
    tests: VbaTestCase[];
    modulesScanned: number;
    modulesIgnored: number;
    contract: string;
}

export interface VbaTestRunItem {
    test: VbaTestCase;
    status: VbaTestStatus;
    durationMs: number;
    error?: string;
}

export interface VbaTestRunReport {
    filePath: string;
    workbookName: string;
    startedAt: string;
    durationMs: number;
    discovery: VbaTestDiscoveryResult;
    results: VbaTestRunItem[];
}

export interface VbaTestRunSummary {
    total: number;
    passed: number;
    failed: number;
    skipped: number;
}

export function discoverVbaTestsFromModule(module: VbaTestModuleEntry): VbaTestCase[] {
    if (module.type !== 'standard' || module.source === undefined) {
        return [];
    }

    const source = module.source;
    const ast = parseModule(source);
    const starts = lineStartOffsets(source);
    const lines = source.split(/\r\n|\r|\n/);
    const tests: VbaTestCase[] = [];

    for (const member of ast.members) {
        if (!isRunnableTestProcedure(member)) {
            continue;
        }
        const annotationLine = precedingTestAnnotationLine(lines, starts, member.span.start);
        if (annotationLine === undefined) {
            continue;
        }
        const location = offsetToLineColumn(starts, member.span.start);
        tests.push({
            id: `${module.name}.${member.name}`,
            moduleName: module.name,
            moduleType: module.type,
            procedureName: member.name,
            qualifiedName: `${module.name}.${member.name}`,
            line: location.line,
            column: location.column,
            annotationLine,
        });
    }

    return tests;
}

export async function discoverWorkbookVbaTests(
    bridge: PythonBridge,
    filePath: string,
): Promise<VbaTestDiscoveryResult> {
    const modules = await bridge.call<Array<{ name: string; type: string }>>(
        'listModules',
        { path: filePath },
    );
    const orderedModules = [...modules].sort(compareVbaModulesForTreeOrder);
    const testableModules = orderedModules.filter((module) => module.type === 'standard');
    const tests: VbaTestCase[] = [];

    for (const module of testableModules) {
        const result = await bridge.call<{ source: string }>(
            'readModule',
            { path: filePath, module: module.name },
        );
        tests.push(...discoverVbaTestsFromModule({
            name: module.name,
            type: module.type,
            source: result.source,
        }));
    }

    return {
        filePath,
        tests,
        modulesScanned: testableModules.length,
        modulesIgnored: orderedModules.length - testableModules.length,
        contract: `Standard-module no-argument Sub procedures with an immediately preceding '${XLIDE_VBA_TEST_DIRECTIVE}' comment directive.`,
    };
}

export function createVbaTestRunReport(input: {
    filePath: string;
    startedAt: Date;
    durationMs: number;
    discovery: VbaTestDiscoveryResult;
    results: readonly VbaTestRunItem[];
}): VbaTestRunReport {
    return {
        filePath: input.filePath,
        workbookName: path.basename(input.filePath),
        startedAt: input.startedAt.toISOString(),
        durationMs: input.durationMs,
        discovery: input.discovery,
        results: [...input.results],
    };
}

export function summarizeVbaTestRun(report: Pick<VbaTestRunReport, 'results'>): VbaTestRunSummary {
    const summary: VbaTestRunSummary = {
        total: report.results.length,
        passed: 0,
        failed: 0,
        skipped: 0,
    };
    for (const result of report.results) {
        summary[result.status]++;
    }
    return summary;
}

export function vbaTestFailureMessage(error: unknown): string {
    const raw = error instanceof Error ? error.message : String(error);
    const pipe = raw.indexOf('|');
    return pipe >= 0 ? raw.slice(pipe + 1) : raw;
}

function isRunnableTestProcedure(member: unknown): member is ProcedureNode {
    return Boolean(
        member &&
        typeof member === 'object' &&
        (member as ProcedureNode).kind === 'Procedure' &&
        (member as ProcedureNode).procKind === 'Sub' &&
        (member as ProcedureNode).params.length === 0,
    );
}

function precedingTestAnnotationLine(lines: readonly string[], starts: readonly number[], offset: number): number | undefined {
    const declarationLine = offsetToLineColumn(starts, offset).line - 1;
    for (let lineIndex = declarationLine - 1; lineIndex >= 0; lineIndex--) {
        const line = lines[lineIndex] ?? '';
        if (/^\s*$/.test(line)) {
            return undefined;
        }
        if (isXlideTestDirectiveComment(line)) {
            return lineIndex + 1;
        }
        if (/^\s*'/.test(line)) {
            continue;
        }
        return undefined;
    }
    return undefined;
}

function isXlideTestDirectiveComment(line: string): boolean {
    return /^\s*'\s*@xlide-test\b/i.test(line);
}

function offsetToLineColumn(
    starts: readonly number[],
    offset: number,
): { line: number; column: number } {
    let lo = 0;
    let hi = starts.length - 1;
    while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (starts[mid] <= offset) { lo = mid; } else { hi = mid - 1; }
    }
    return { line: lo + 1, column: offset - starts[lo] + 1 };
}
