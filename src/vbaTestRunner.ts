import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import { parseModule } from './analyzer/parser/parseModule';
import type { ModuleMember, ModuleNode, ProcedureNode, Span } from './analyzer/parser/nodes';
import { lineStartOffsets } from './vbaSourceScan';
import { compareVbaModulesForTreeOrder } from './moduleDisplay';
import { measurePerformance } from './performanceTrace';
import { isReadModulesUnavailable } from './pythonBridgeErrors';

export const XLIDE_VBA_TEST_DIRECTIVE = '@xlide-test';
export const VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE = 'vba-test-directive';

export type VbaTestStatus =
    | 'passed'
    | 'failed'
    | 'skipped'
    | 'xfail'
    | 'xpass'
    | 'timeout'
    | 'host-error';

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
    expectedError?: number | 'any';
    skipReason?: string;
    xfailReason?: string;
}

export interface VbaTestDirectiveIssue {
    code: typeof VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE;
    message: string;
    span: Span;
    line: number;
    column: number;
}

export interface VbaTestSelectionOptions {
    moduleName?: string;
    procedureName?: string;
    testIds?: readonly string[];
    includeTags?: readonly string[];
    excludeTags?: readonly string[];
}

export interface VbaTestTagSummary {
    name: string;
    testCount: number;
}

export interface VbaTestDiscoveryResult {
    filePath: string;
    tests: VbaTestCase[];
    unfilteredTestCount: number;
    selection?: VbaTestSelectionOptions;
    modulesScanned: number;
    modulesIgnored: number;
    contract: string;
}

export interface VbaTestRunItem {
    test: VbaTestCase;
    status: VbaTestStatus;
    durationMs: number;
    error?: string;
    output?: string[];
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
    timeout: number;
    hostError: number;
}

export function discoverVbaTestsFromModule(
    module: VbaTestModuleEntry,
    parsedModule?: ModuleNode,
): VbaTestCase[] {
    if (module.type !== 'standard' || module.source === undefined) {
        return [];
    }

    const source = module.source;
    const ast = parsedModule ?? parseModule(source);
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

export function validateVbaTestDirectivesFromModule(
    module: VbaTestModuleEntry,
    parsedModule?: ModuleNode,
): VbaTestDirectiveIssue[] {
    if (module.source === undefined) {
        return [];
    }

    const source = module.source;
    const ast = parsedModule ?? parseModule(source);
    const starts = lineStartOffsets(source);
    const lines = source.split(/\r\n|\r|\n/);
    const issues: VbaTestDirectiveIssue[] = [];
    const candidatesByLine = testDirectiveCandidatesByLine(lines, starts);
    const attachedLines = new Set<number>();

    for (const member of ast.members) {
        const block = precedingCommentBlock(lines, starts, member.span.start);
        const blockCandidates = block
            .map((entry) => candidatesByLine.get(entry.line))
            .filter((candidate): candidate is TestDirectiveCandidate => candidate !== undefined);
        if (blockCandidates.length === 0) {
            continue;
        }

        const targetIssue = testDirectiveTargetIssue(module, member);
        for (const candidate of blockCandidates) {
            attachedLines.add(candidate.line);
            const parsed = parseXlideTestDirective(candidate.text);
            if (!parsed) {
                issues.push(testDirectiveIssue(candidate, unknownTestDirectiveMessage()));
                continue;
            }
            issues.push(...validateTestDirectiveMetadata(candidate, parsed));
            if (targetIssue) {
                issues.push(testDirectiveIssue(candidate, targetIssue));
            }
        }
    }

    for (const candidate of candidatesByLine.values()) {
        if (attachedLines.has(candidate.line)) {
            continue;
        }
        const parsed = parseXlideTestDirective(candidate.text);
        if (!parsed) {
            issues.push(testDirectiveIssue(candidate, unknownTestDirectiveMessage()));
            continue;
        }
        issues.push(...validateTestDirectiveMetadata(candidate, parsed));
        issues.push(testDirectiveIssue(
            candidate,
            'XLIDE test directives must be in the comment block immediately above a zero-argument Sub procedure.',
        ));
    }

    return issues.sort((left, right) =>
        left.span.start - right.span.start || left.message.localeCompare(right.message),
    );
}

// Batch read of every module in one workbook open; falls back to listModules
// plus one readModule per standard module for backends without readModules.
async function listWorkbookModulesForDiscovery(
    bridge: PythonBridge,
    filePath: string,
): Promise<VbaTestModuleEntry[]> {
    try {
        const modules = await bridge.call<VbaTestModuleEntry[]>(
            'readModules',
            { path: filePath },
        );
        return modules.filter((module) => typeof module.source === 'string');
    } catch (err) {
        if (!isReadModulesUnavailable(err)) {
            throw err;
        }
    }
    return bridge.call<VbaTestModuleEntry[]>('listModules', { path: filePath });
}

export async function discoverWorkbookVbaTests(
    bridge: PythonBridge,
    filePath: string,
    selection?: VbaTestSelectionOptions,
): Promise<VbaTestDiscoveryResult> {
    return measurePerformance('vbaTests.discoverWorkbook', path.basename(filePath), async () => {
    const normalizedSelection = normalizeVbaTestSelection(selection);
    const modules = await listWorkbookModulesForDiscovery(bridge, filePath);
    const orderedModules = [...modules].sort(compareVbaModulesForTreeOrder);
    const testableModules = orderedModules.filter((module) =>
        module.type === 'standard' &&
        (!normalizedSelection?.moduleName || equalsIgnoreCase(module.name, normalizedSelection.moduleName)),
    );
    const discoveredTests: VbaTestCase[] = [];

    for (const module of testableModules) {
        const source = module.source ?? (await bridge.call<{ source: string }>(
            'readModule',
            { path: filePath, module: module.name },
        )).source;
        discoveredTests.push(...discoverVbaTestsFromModule({
            name: module.name,
            type: module.type,
            source,
        }));
    }
    const tests = filterVbaTests(discoveredTests, normalizedSelection);

    return {
        filePath,
        tests,
        unfilteredTestCount: discoveredTests.length,
        selection: normalizedSelection,
        modulesScanned: testableModules.length,
        modulesIgnored: orderedModules.length - testableModules.length,
        contract: `Standard-module no-argument Sub procedures with an immediately preceding '${XLIDE_VBA_TEST_DIRECTIVE}' comment directive.`,
    };
    });
}

export function filterVbaTests(
    tests: readonly VbaTestCase[],
    selection?: VbaTestSelectionOptions,
): VbaTestCase[] {
    const normalizedSelection = normalizeVbaTestSelection(selection);
    if (!normalizedSelection) {
        return [...tests];
    }

    const includeTags = normalizedSelection.includeTags?.map(normalizeTag) ?? [];
    const excludeTags = normalizedSelection.excludeTags?.map(normalizeTag) ?? [];
    const testIds = normalizedSelection.testIds
        ? new Set(normalizedSelection.testIds.map(normalizeTestId))
        : undefined;

    return tests.filter((test) => {
        if (testIds && !testIds.has(normalizeTestId(test.id))) {
            return false;
        }
        if (normalizedSelection.moduleName && !equalsIgnoreCase(test.moduleName, normalizedSelection.moduleName)) {
            return false;
        }
        if (
            normalizedSelection.procedureName &&
            !equalsIgnoreCase(test.procedureName, normalizedSelection.procedureName)
        ) {
            return false;
        }

        const tags = new Set(test.metadata.tags.map(normalizeTag));
        if (includeTags.length > 0 && !includeTags.some((tag) => tags.has(tag))) {
            return false;
        }
        if (excludeTags.length > 0 && excludeTags.some((tag) => tags.has(tag))) {
            return false;
        }
        return true;
    });
}

export function summarizeVbaTestTags(tests: readonly VbaTestCase[]): VbaTestTagSummary[] {
    const counts = new Map<string, Set<string>>();
    for (const test of tests) {
        for (const tag of test.metadata.tags) {
            const normalized = normalizeTag(tag);
            if (!normalized) {
                continue;
            }
            const bucket = counts.get(normalized) ?? new Set<string>();
            bucket.add(test.id);
            counts.set(normalized, bucket);
        }
    }
    return [...counts.entries()]
        .map(([name, testIds]) => ({ name, testCount: testIds.size }))
        .sort((left, right) => left.name.localeCompare(right.name, undefined, { sensitivity: 'base' }));
}

export function describeVbaTestSelection(selection?: VbaTestSelectionOptions): string {
    const normalizedSelection = normalizeVbaTestSelection(selection);
    if (!normalizedSelection) {
        return '';
    }

    const parts: string[] = [];
    if (normalizedSelection.moduleName) {
        parts.push(`module ${normalizedSelection.moduleName}`);
    }
    if (normalizedSelection.procedureName) {
        parts.push(`test ${normalizedSelection.procedureName}`);
    }
    if (normalizedSelection.testIds?.length) {
        parts.push(`${normalizedSelection.testIds.length} selected test${normalizedSelection.testIds.length === 1 ? '' : 's'}`);
    }
    if (normalizedSelection.includeTags?.length) {
        parts.push(`tags ${normalizedSelection.includeTags.join(', ')}`);
    }
    if (normalizedSelection.excludeTags?.length) {
        parts.push(`excluding ${normalizedSelection.excludeTags.join(', ')}`);
    }
    return parts.join(', ');
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
        timeout: 0,
        hostError: 0,
    };
    for (const result of report.results) {
        if (result.status === 'host-error') {
            summary.hostError++;
        } else {
            summary[result.status]++;
        }
    }
    return summary;
}

function normalizeVbaTestSelection(selection?: VbaTestSelectionOptions): VbaTestSelectionOptions | undefined {
    const moduleName = normalizeOptionalText(selection?.moduleName);
    const procedureName = normalizeOptionalText(selection?.procedureName);
    const testIds = normalizeTestIdList(selection?.testIds);
    const includeTags = normalizeTagList(selection?.includeTags);
    const excludeTags = normalizeTagList(selection?.excludeTags);
    if (!moduleName && !procedureName && testIds.length === 0 && includeTags.length === 0 && excludeTags.length === 0) {
        return undefined;
    }
    return {
        ...(moduleName ? { moduleName } : {}),
        ...(procedureName ? { procedureName } : {}),
        ...(testIds.length > 0 ? { testIds } : {}),
        ...(includeTags.length > 0 ? { includeTags } : {}),
        ...(excludeTags.length > 0 ? { excludeTags } : {}),
    };
}

function normalizeOptionalText(value: string | undefined): string | undefined {
    const trimmed = value?.trim();
    return trimmed ? trimmed : undefined;
}

function normalizeTagList(values: readonly string[] | undefined): string[] {
    return [...new Set((values ?? [])
        .map((value) => value.trim())
        .filter((value) => value.length > 0))];
}

function normalizeTestIdList(values: readonly string[] | undefined): string[] {
    return [...new Set((values ?? [])
        .map((value) => value.trim())
        .filter((value) => value.length > 0))];
}

function normalizeTestId(value: string): string {
    return value.trim().toLowerCase();
}

function normalizeTag(value: string): string {
    return value.trim().toLowerCase();
}

function equalsIgnoreCase(left: string, right: string): boolean {
    return left.localeCompare(right, undefined, { sensitivity: 'accent' }) === 0;
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
    const block = precedingCommentBlock(lines, starts, offset);
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

function precedingCommentBlock(
    lines: readonly string[],
    starts: readonly number[],
    offset: number,
): Array<{ line: number; text: string }> {
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
    return block;
}

interface ParsedTestDirective {
    kind: 'test' | 'skip' | 'xfail';
    values: Record<string, string>;
    metadata: ParsedDirectiveKeyValues;
}

interface ParsedDirectiveKeyValues {
    values: Record<string, string>;
    entries: Array<{ key: string; value: string }>;
    malformedSegments: string[];
}

interface TestDirectiveCandidate {
    line: number;
    column: number;
    text: string;
    span: Span;
}

function parseXlideTestDirective(line: string): ParsedTestDirective | undefined {
    const match = /^\s*'\s*@(xlide-test(?:-(skip|xfail))?)\b(.*)$/i.exec(line);
    if (!match) {
        return undefined;
    }
    const suffix = match[2]?.toLowerCase();
    const metadata = parseDirectiveKeyValues(match[3] ?? '');
    return {
        kind: suffix === 'skip' || suffix === 'xfail' ? suffix : 'test',
        values: metadata.values,
        metadata,
    };
}

function parseDirectiveKeyValues(text: string): ParsedDirectiveKeyValues {
    const values: Record<string, string> = {};
    const entries: Array<{ key: string; value: string }> = [];
    const malformedSegments: string[] = [];
    const metadataText = stripInlineMetadataComment(text);
    const re = /([A-Za-z][A-Za-z0-9_-]*)=(?:"([^"]*)"|'([^']*)'|(\S+))/g;
    let match: RegExpExecArray | null;
    let cursor = 0;
    while ((match = re.exec(metadataText)) !== null) {
        const gap = metadataText.slice(cursor, match.index).trim();
        if (gap.length > 0) {
            consumeStandaloneMetadata(gap, values, entries, malformedSegments);
        }
        const key = match[1].toLowerCase();
        const value = match[2] ?? match[3] ?? match[4] ?? '';
        values[key] = value;
        entries.push({ key, value });
        cursor = match.index + match[0].length;
    }
    const tail = metadataText.slice(cursor).trim();
    if (tail.length > 0) {
        consumeStandaloneMetadata(tail, values, entries, malformedSegments);
    }
    return { values, entries, malformedSegments };
}

function consumeStandaloneMetadata(
    text: string,
    values: Record<string, string>,
    entries: Array<{ key: string; value: string }>,
    malformedSegments: string[],
): void {
    const malformed: string[] = [];
    for (const token of text.split(/\s+/).filter(Boolean)) {
        const key = token.toLowerCase();
        if (key === 'expected-error' || key === 'expectederror') {
            values['expected-error'] = 'any';
            entries.push({ key: 'expected-error', value: 'any' });
        } else {
            malformed.push(token);
        }
    }
    if (malformed.length > 0) {
        malformedSegments.push(malformed.join(' '));
    }
}

function stripInlineMetadataComment(text: string): string {
    const commentStart = text.search(/\s--/);
    return commentStart >= 0 ? text.slice(0, commentStart) : text;
}

function testDirectiveCandidatesByLine(
    lines: readonly string[],
    starts: readonly number[],
): Map<number, TestDirectiveCandidate> {
    const candidates = new Map<number, TestDirectiveCandidate>();
    lines.forEach((line, lineIndex) => {
        const candidate = testDirectiveCandidateForLine(line, lineIndex, starts);
        if (candidate) {
            candidates.set(candidate.line, candidate);
        }
    });
    return candidates;
}

function testDirectiveCandidateForLine(
    line: string,
    lineIndex: number,
    starts: readonly number[],
): TestDirectiveCandidate | undefined {
    const commentMatch = /^(\s*)'/.exec(line);
    if (!commentMatch) {
        return undefined;
    }
    const apostropheIndex = commentMatch[1].length;
    if (line.startsWith("'''", apostropheIndex)) {
        return undefined;
    }
    const body = line.slice(apostropheIndex + 1).trimStart();
    const token = /^@\S*/.exec(body)?.[0] ?? '';
    if (!isLikelyXlideTestDirectiveToken(token)) {
        return undefined;
    }
    const atIndex = line.indexOf('@', apostropheIndex);
    const spanStart = (starts[lineIndex] ?? 0) + Math.max(0, atIndex);
    const spanEnd = Math.max(spanStart + 1, (starts[lineIndex] ?? 0) + line.length);
    return {
        line: lineIndex + 1,
        column: Math.max(1, atIndex + 1),
        text: line,
        span: { start: spanStart, end: spanEnd },
    };
}

function isLikelyXlideTestDirectiveToken(token: string): boolean {
    const normalized = token.trim().toLowerCase();
    if (normalized === XLIDE_VBA_TEST_DIRECTIVE || normalized.startsWith(`${XLIDE_VBA_TEST_DIRECTIVE}-`)) {
        return true;
    }
    return normalized.startsWith('@xlide-') &&
        normalized.length <= XLIDE_VBA_TEST_DIRECTIVE.length + 4 &&
        editDistance(normalized, XLIDE_VBA_TEST_DIRECTIVE) <= 2;
}

function editDistance(left: string, right: string): number {
    const previous = Array.from({ length: right.length + 1 }, (_, index) => index);
    const current = new Array<number>(right.length + 1);
    for (let i = 1; i <= left.length; i++) {
        current[0] = i;
        for (let j = 1; j <= right.length; j++) {
            const cost = left[i - 1] === right[j - 1] ? 0 : 1;
            current[j] = Math.min(
                current[j - 1] + 1,
                previous[j] + 1,
                previous[j - 1] + cost,
            );
        }
        previous.splice(0, previous.length, ...current);
    }
    return previous[right.length];
}

function validateTestDirectiveMetadata(
    candidate: TestDirectiveCandidate,
    directive: ParsedTestDirective,
): VbaTestDirectiveIssue[] {
    const issues: VbaTestDirectiveIssue[] = [];
    const supportedKeys = directive.kind === 'test'
        ? TEST_DIRECTIVE_METADATA_KEYS
        : TEST_DIRECTIVE_METADATA_KEYS_WITH_REASON;

    if (directive.metadata.malformedSegments.length > 0) {
        issues.push(testDirectiveIssue(
            candidate,
            'Malformed XLIDE test metadata. Use key=value pairs and quote values that contain spaces.',
        ));
    }

    for (const entry of directive.metadata.entries) {
        if (!supportedKeys.has(entry.key)) {
            issues.push(testDirectiveIssue(
                candidate,
                `Unknown XLIDE test metadata key '${entry.key}'.`,
            ));
        }
    }

    if (hasMetadataKey(directive.values, 'tags') && splitTags(directive.values.tags).length === 0) {
        issues.push(testDirectiveIssue(candidate, 'XLIDE test metadata key tags must list at least one tag.'));
    }

    const timeoutValue = directive.values.timeout ?? directive.values.timeoutms;
    if (
        (hasMetadataKey(directive.values, 'timeout') || hasMetadataKey(directive.values, 'timeoutms')) &&
        parseTimeoutMs(timeoutValue) === undefined
    ) {
        issues.push(testDirectiveIssue(candidate, 'XLIDE test timeout must be a positive integer with optional ms or s suffix.'));
    }

    const expectedErrorValue = directive.values['expected-error'] ?? directive.values.expectederror;
    if (
        (hasMetadataKey(directive.values, 'expected-error') || hasMetadataKey(directive.values, 'expectederror')) &&
        parseExpectedVbaErrorMetadata(expectedErrorValue) === undefined
    ) {
        issues.push(testDirectiveIssue(candidate, 'XLIDE expected-error metadata must be a positive VBA error number or any.'));
    }

    if (
        (directive.kind === 'skip' || directive.kind === 'xfail') &&
        !directive.values.reason?.trim()
    ) {
        issues.push(testDirectiveIssue(
            candidate,
            `XLIDE ${directive.kind === 'skip' ? 'skip' : 'expected-failure'} test directives should include reason="...".`,
        ));
    }

    return issues;
}

const TEST_DIRECTIVE_METADATA_KEYS = new Set([
    'tags',
    'owner',
    'requirement',
    'req',
    'timeout',
    'timeoutms',
    'expected-error',
    'expectederror',
]);

const TEST_DIRECTIVE_METADATA_KEYS_WITH_REASON = new Set([
    ...TEST_DIRECTIVE_METADATA_KEYS,
    'reason',
]);

function hasMetadataKey(values: Record<string, string>, key: string): boolean {
    return Object.prototype.hasOwnProperty.call(values, key);
}

function testDirectiveTargetIssue(module: VbaTestModuleEntry, member: ModuleMember): string | undefined {
    const moduleType = module.type || 'unknown';
    if (moduleType !== 'standard') {
        return `XLIDE test directives only run from standard modules; '${module.name}' is a ${moduleType} module.`;
    }
    if (!isProcedureNode(member)) {
        return 'XLIDE test directives must target a zero-argument Sub procedure.';
    }
    if (member.procKind !== 'Sub') {
        return 'XLIDE test directives must target a Sub procedure; Functions and Properties are not runnable tests.';
    }
    if (member.params.length > 0) {
        return 'XLIDE test Sub procedures must not declare parameters.';
    }
    return undefined;
}

function isProcedureNode(member: ModuleMember): member is ProcedureNode {
    return member.kind === 'Procedure';
}

function testDirectiveIssue(candidate: TestDirectiveCandidate, message: string): VbaTestDirectiveIssue {
    return {
        code: VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE,
        message,
        span: candidate.span,
        line: candidate.line,
        column: candidate.column,
    };
}

function unknownTestDirectiveMessage(): string {
    return 'Unknown XLIDE test directive. Supported directives are @xlide-test, @xlide-test-skip, and @xlide-test-xfail.';
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
    const expectedError = parseExpectedVbaErrorMetadata(directive.values['expected-error'] ?? directive.values.expectederror);
    metadata.expectedError = expectedError ?? metadata.expectedError;
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

function parseExpectedVbaErrorNumber(value: string | undefined): number | undefined {
    const normalized = value?.trim();
    if (!normalized || !/^\d+$/.test(normalized)) {
        return undefined;
    }
    const amount = Number(normalized);
    return Number.isSafeInteger(amount) && amount > 0 ? amount : undefined;
}

function parseExpectedVbaErrorMetadata(value: string | undefined): number | 'any' | undefined {
    const normalized = value?.trim().toLowerCase();
    if (normalized === 'any') {
        return 'any';
    }
    return parseExpectedVbaErrorNumber(value);
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
