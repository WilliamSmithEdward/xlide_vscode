import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import { parseModule } from './analyzer/parser/parseModule';
import type { ProcedureNode } from './analyzer/parser/nodes';
import { lineStartOffsets } from './vbaStructuralAnalysis';
import { compareVbaModulesForTreeOrder } from './moduleDisplay';

export const XLIDE_VBA_TEST_DIRECTIVE = '@xlide-test';

export type VbaTestStatus = 'passed' | 'failed' | 'skipped' | 'xfail' | 'xpass';

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
    metadata: VbaTestMetadata;
}

export interface VbaTestMetadata {
    tags: string[];
    owner?: string;
    requirement?: string;
    timeoutMs?: number;
    expectedError?: string;
    skipReason?: string;
    xfailReason?: string;
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
    xfail: number;
    xpass: number;
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
        const annotation = precedingTestAnnotation(lines, starts, member.span.start);
        if (!annotation) {
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
            annotationLine: annotation.line,
            metadata: annotation.metadata,
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
        xfail: 0,
        xpass: 0,
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

function precedingTestAnnotation(
    lines: readonly string[],
    starts: readonly number[],
    offset: number,
): { line: number; metadata: VbaTestMetadata } | undefined {
    const declarationLine = offsetToLineColumn(starts, offset).line - 1;
    const block: Array<{ line: number; text: string }> = [];
    for (let lineIndex = declarationLine - 1; lineIndex >= 0; lineIndex--) {
        const line = lines[lineIndex] ?? '';
        if (/^\s*$/.test(line)) {
            break;
        }
        if (!/^\s*'/.test(line)) {
            break;
        }
        block.unshift({ line: lineIndex + 1, text: line });
    }

    const metadata = defaultTestMetadata();
    let annotationLine: number | undefined;
    let discovered = false;
    for (const entry of block) {
        const directive = parseXlideTestDirective(entry.text);
        if (!directive) {
            continue;
        }
        if (annotationLine === undefined) {
            annotationLine = entry.line;
        }
        discovered = true;
        mergeDirectiveMetadata(metadata, directive);
    }

    if (!discovered || annotationLine === undefined) {
        return undefined;
    }
    return {
        line: annotationLine,
        metadata,
    };
}

interface ParsedTestDirective {
    kind: 'test' | 'skip' | 'xfail';
    values: Record<string, string>;
}

function parseXlideTestDirective(line: string): ParsedTestDirective | undefined {
    const match = /^\s*'\s*@(xlide-test(?:-(skip|xfail))?)\b(.*)$/i.exec(line);
    if (!match) {
        return undefined;
    }
    const suffix = match[2]?.toLowerCase();
    return {
        kind: suffix === 'skip' || suffix === 'xfail' ? suffix : 'test',
        values: parseDirectiveKeyValues(match[3] ?? ''),
    };
}

function parseDirectiveKeyValues(text: string): Record<string, string> {
    const values: Record<string, string> = {};
    const re = /([A-Za-z][A-Za-z0-9_-]*)=(?:"([^"]*)"|'([^']*)'|(\S+))/g;
    let match: RegExpExecArray | null;
    while ((match = re.exec(text)) !== null) {
        const key = match[1].toLowerCase();
        values[key] = match[2] ?? match[3] ?? match[4] ?? '';
    }
    return values;
}

function defaultTestMetadata(): VbaTestMetadata {
    return { tags: [] };
}

function mergeDirectiveMetadata(metadata: VbaTestMetadata, directive: ParsedTestDirective): void {
    const tags = splitTags(directive.values.tags);
    if (tags.length > 0) {
        metadata.tags = [...new Set([...metadata.tags, ...tags])];
    }
    metadata.owner = directive.values.owner ?? metadata.owner;
    metadata.requirement = directive.values.requirement ?? directive.values.req ?? metadata.requirement;
    metadata.expectedError = directive.values['expected-error'] ?? directive.values.expectederror ?? metadata.expectedError;
    const timeoutMs = parseTimeoutMs(directive.values.timeout ?? directive.values.timeoutms);
    if (timeoutMs !== undefined) {
        metadata.timeoutMs = timeoutMs;
    }
    if (directive.kind === 'skip') {
        metadata.skipReason = directive.values.reason ?? 'Skipped by @xlide-test-skip.';
    } else if (directive.kind === 'xfail') {
        metadata.xfailReason = directive.values.reason ?? 'Expected failure by @xlide-test-xfail.';
    }
}

function splitTags(value: string | undefined): string[] {
    if (!value) {
        return [];
    }
    return value
        .split(',')
        .map((tag) => tag.trim())
        .filter((tag) => tag.length > 0);
}

function parseTimeoutMs(value: string | undefined): number | undefined {
    if (!value) {
        return undefined;
    }
    const normalized = value.trim().toLowerCase();
    const match = /^(\d+)(ms|s)?$/.exec(normalized);
    if (!match) {
        return undefined;
    }
    const amount = Number(match[1]);
    if (!Number.isFinite(amount) || amount <= 0) {
        return undefined;
    }
    return match[2] === 's' ? amount * 1000 : amount;
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
