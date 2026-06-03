export interface VbaTestDirectiveCompletionRange {
    start: number;
    end: number;
}

export interface VbaTestDirectiveCompletion {
    label: string;
    insertText: string;
    detail: string;
    documentation: string;
    sortText: string;
    range: VbaTestDirectiveCompletionRange;
    exclusive: boolean;
}

interface TestDirectiveCompletionTemplate {
    label: string;
    directiveText: string;
    detail: string;
    documentation: string;
    sortText: string;
}

interface TestDirectiveMetadataCompletionTemplate {
    label: string;
    insertText: string;
    detail: string;
    documentation: string;
    sortText: string;
    canonicalKey: TestDirectiveMetadataKey;
}

interface TestDirectiveValueCompletionTemplate {
    label: string;
    insertText: string;
    detail: string;
    documentation: string;
    sortText: string;
}

type TestDirectiveKind = 'test' | 'skip' | 'xfail';
type TestDirectiveMetadataKey = 'tags' | 'owner' | 'requirement' | 'timeout' | 'expected-error' | 'reason';

const TEST_DIRECTIVE_COMPLETIONS: readonly TestDirectiveCompletionTemplate[] = [
    {
        label: '@xlide-test',
        directiveText: '@xlide-test',
        detail: 'XLIDE VBA test',
        documentation: 'Marks the following no-argument standard-module Sub as an XLIDE test.',
        sortText: '0:@xlide-test',
    },
    {
        label: '@xlide-test-skip',
        directiveText: '@xlide-test-skip reason="$1"',
        detail: 'XLIDE skipped VBA test',
        documentation: 'Marks the following XLIDE test as skipped and records the reason.',
        sortText: '1:@xlide-test-skip',
    },
    {
        label: '@xlide-test-xfail',
        directiveText: '@xlide-test-xfail reason="$1"',
        detail: 'XLIDE expected-failure VBA test',
        documentation: 'Marks the following XLIDE test as expected to fail and records the reason.',
        sortText: '2:@xlide-test-xfail',
    },
];

const DIRECTIVE_TOKEN_RE = /^@?[A-Za-z0-9-]*$/;
const DIRECTIVE_SUFFIX_RE = /[A-Za-z0-9-]/;
const METADATA_KEY_RE = /^[A-Za-z][A-Za-z0-9_-]*$/;
const METADATA_KEY_SUFFIX_RE = /[A-Za-z0-9_-]/;
const METADATA_VALUE_SUFFIX_RE = /[^\s]/;

export function resolveVbaTestDirectiveCompletions(
    line: string,
    character: number,
): VbaTestDirectiveCompletion[] {
    const cursor = Math.max(0, Math.min(character, line.length));
    const before = line.slice(0, cursor);
    const comment = /^\s*'(?!'')\s*/.exec(before);
    const bare = comment ? undefined : /^\s*/.exec(before);
    if (!comment && !bare) {
        return [];
    }

    const inComment = Boolean(comment);
    const start = comment ? comment[0].length : bare![0].length;
    const typed = before.slice(start);
    if (DIRECTIVE_TOKEN_RE.test(typed)) {
        return resolveDirectiveNameCompletions(line, cursor, start, typed, inComment);
    }

    if (!inComment) {
        return [];
    }

    return resolveDirectiveMetadataCompletions(line, cursor, start);
}

function resolveDirectiveNameCompletions(
    line: string,
    cursor: number,
    start: number,
    typed: string,
    inComment: boolean,
): VbaTestDirectiveCompletion[] {
    const normalized = typed.startsWith('@')
        ? typed.toLowerCase()
        : `@${typed.toLowerCase()}`;
    const matches = TEST_DIRECTIVE_COMPLETIONS.filter((completion) => {
        const label = completion.label.toLowerCase();
        return label === normalized || label.startsWith(normalized);
    });
    if (matches.length === 0) {
        return [];
    }

    let end = cursor;
    while (end < line.length && DIRECTIVE_SUFFIX_RE.test(line[end])) {
        end += 1;
    }

    return matches.map((completion) => ({
        label: completion.label,
        insertText: inComment ? completion.directiveText : `' ${completion.directiveText}`,
        detail: completion.detail,
        documentation: completion.documentation,
        sortText: completion.sortText,
        range: { start, end },
        exclusive: inComment,
    }));
}

function resolveDirectiveMetadataCompletions(
    line: string,
    cursor: number,
    commentContentStart: number,
): VbaTestDirectiveCompletion[] {
    const commentTextBefore = line.slice(commentContentStart, cursor);
    const directive = directiveAtCommentStart(commentTextBefore);
    if (!directive) {
        return [];
    }

    const metadataBefore = commentTextBefore.slice(directive.end);
    if (/\s--/.test(metadataBefore)) {
        return [];
    }

    const tokenStart = currentMetadataTokenStart(line, cursor, commentContentStart + directive.end);
    const token = line.slice(tokenStart, cursor);
    const equalsIndex = token.indexOf('=');
    if (equalsIndex >= 0) {
        return resolveMetadataValueCompletions(line, cursor, tokenStart, token, directive.kind);
    }

    if (isInsideMetadataQuote(metadataBefore)) {
        return [];
    }

    if (!METADATA_KEY_RE.test(token) && token.length > 0) {
        return [];
    }

    const usedKeys = usedMetadataKeys(line.slice(commentContentStart + directive.end));
    const typedKey = token.toLowerCase();
    const matches = metadataCompletionTemplates(directive.kind)
        .filter((completion) => !usedKeys.has(completion.canonicalKey))
        .filter((completion) => completion.label.toLowerCase().startsWith(typedKey));
    if (matches.length === 0) {
        return [];
    }

    let end = cursor;
    while (end < line.length && METADATA_KEY_SUFFIX_RE.test(line[end])) {
        end += 1;
    }

    return matches.map((completion) => ({
        label: completion.label,
        insertText: completion.insertText,
        detail: completion.detail,
        documentation: completion.documentation,
        sortText: completion.sortText,
        range: { start: tokenStart, end },
        exclusive: true,
    }));
}

function resolveMetadataValueCompletions(
    line: string,
    cursor: number,
    tokenStart: number,
    token: string,
    directiveKind: TestDirectiveKind,
): VbaTestDirectiveCompletion[] {
    const equalsIndex = token.indexOf('=');
    const rawKey = token.slice(0, equalsIndex);
    const key = canonicalMetadataKey(rawKey);
    if (!key || !metadataKeyIsSupportedForDirective(key, directiveKind)) {
        return [];
    }

    const valueStart = tokenStart + equalsIndex + 1;
    const typedValue = line.slice(valueStart, cursor).toLowerCase();
    const matches = valueCompletionTemplates(key, directiveKind)
        .filter((completion) => completion.label.toLowerCase().startsWith(typedValue));
    if (matches.length === 0) {
        return [];
    }

    let end = cursor;
    while (end < line.length && METADATA_VALUE_SUFFIX_RE.test(line[end])) {
        end += 1;
    }

    return matches.map((completion) => ({
        label: completion.label,
        insertText: completion.insertText,
        detail: completion.detail,
        documentation: completion.documentation,
        sortText: completion.sortText,
        range: { start: valueStart, end },
        exclusive: true,
    }));
}

function directiveAtCommentStart(text: string): { kind: TestDirectiveKind; end: number } | undefined {
    const match = /^@(xlide-test(?:-(skip|xfail))?)\b/i.exec(text);
    if (!match) {
        return undefined;
    }
    const suffix = match[2]?.toLowerCase();
    return {
        kind: suffix === 'skip' || suffix === 'xfail' ? suffix : 'test',
        end: match[0].length,
    };
}

function metadataCompletionTemplates(kind: TestDirectiveKind): TestDirectiveMetadataCompletionTemplate[] {
    const common: TestDirectiveMetadataCompletionTemplate[] = [
        {
            label: 'tags=',
            insertText: 'tags=${1:smoke,fast}',
            detail: 'XLIDE test metadata',
            documentation: 'Adds comma-separated tags used by filtered test runs.',
            sortText: '1:tags',
            canonicalKey: 'tags',
        },
        {
            label: 'owner=',
            insertText: 'owner=${1:owner}',
            detail: 'XLIDE test metadata',
            documentation: 'Records the owning team or person for this test.',
            sortText: '2:owner',
            canonicalKey: 'owner',
        },
        {
            label: 'requirement=',
            insertText: 'requirement=${1:REQ-001}',
            detail: 'XLIDE test metadata',
            documentation: 'Links this test to a requirement or tracking identifier.',
            sortText: '3:requirement',
            canonicalKey: 'requirement',
        },
        {
            label: 'timeout=',
            insertText: 'timeout=${1:10s}',
            detail: 'XLIDE test metadata',
            documentation: 'Sets the per-test timeout. Use a positive integer with optional ms or s suffix.',
            sortText: '4:timeout',
            canonicalKey: 'timeout',
        },
        {
            label: 'expected-error=',
            insertText: 'expected-error=${1:13}',
            detail: 'XLIDE test metadata',
            documentation: 'Records the expected VBA error number for tests that intentionally exercise an error path.',
            sortText: '5:expected-error',
            canonicalKey: 'expected-error',
        },
    ];

    if (kind === 'test') {
        return common;
    }

    const reasonPlaceholder = kind === 'skip'
        ? 'Requires external workbook'
        : 'Known issue pending fix';
    return [
        {
            label: 'reason=',
            insertText: `reason="\${1:${reasonPlaceholder}}"`,
            detail: 'XLIDE test metadata',
            documentation: 'Records why this test is skipped or expected to fail.',
            sortText: '0:reason',
            canonicalKey: 'reason',
        },
        ...common,
    ];
}

function valueCompletionTemplates(
    key: TestDirectiveMetadataKey,
    directiveKind: TestDirectiveKind,
): TestDirectiveValueCompletionTemplate[] {
    switch (key) {
        case 'timeout':
            return [
                valueCompletion('10s', '10s', 'XLIDE test timeout', 'Run this test with a 10 second timeout.', '0:10s'),
                valueCompletion('30s', '30s', 'XLIDE test timeout', 'Run this test with a 30 second timeout.', '1:30s'),
                valueCompletion('2500ms', '2500ms', 'XLIDE test timeout', 'Run this test with a 2500 millisecond timeout.', '2:2500ms'),
            ];
        case 'reason': {
            const placeholder = directiveKind === 'skip'
                ? 'Requires external workbook'
                : 'Known issue pending fix';
            return [
                valueCompletion(
                    `"${placeholder}"`,
                    `"\${1:${placeholder}}"`,
                    'XLIDE test reason',
                    'Quote reason values when they contain spaces.',
                    '0:reason',
                ),
            ];
        }
        case 'tags':
            return [valueCompletion('smoke,fast', '${1:smoke},${2:fast}', 'XLIDE test tags', 'Use comma-separated tag names.', '0:tags')];
        case 'owner':
            return [valueCompletion('owner', '${1:owner}', 'XLIDE test owner', 'Use a team, person, or ownership token.', '0:owner')];
        case 'requirement':
            return [valueCompletion('REQ-001', '${1:REQ-001}', 'XLIDE test requirement', 'Use a requirement or tracking identifier.', '0:requirement')];
        case 'expected-error':
            return [valueCompletion('13', '${1:13}', 'XLIDE expected error', 'Use the expected VBA error number.', '0:expected-error')];
    }
}

function valueCompletion(
    label: string,
    insertText: string,
    detail: string,
    documentation: string,
    sortText: string,
): TestDirectiveValueCompletionTemplate {
    return { label, insertText, detail, documentation, sortText };
}

function currentMetadataTokenStart(line: string, cursor: number, metadataStart: number): number {
    let start = cursor;
    while (start > metadataStart && !/\s/.test(line[start - 1])) {
        start -= 1;
    }
    return start;
}

function usedMetadataKeys(metadataText: string): Set<TestDirectiveMetadataKey> {
    const keys = new Set<TestDirectiveMetadataKey>();
    const re = /([A-Za-z][A-Za-z0-9_-]*)=(?:"[^"]*"|'[^']*'|\S*)/g;
    let match: RegExpExecArray | null;
    while ((match = re.exec(metadataText)) !== null) {
        const canonical = canonicalMetadataKey(match[1]);
        if (canonical) {
            keys.add(canonical);
        }
    }
    return keys;
}

function canonicalMetadataKey(key: string): TestDirectiveMetadataKey | undefined {
    switch (key.toLowerCase()) {
        case 'tags':
            return 'tags';
        case 'owner':
            return 'owner';
        case 'requirement':
        case 'req':
            return 'requirement';
        case 'timeout':
        case 'timeoutms':
            return 'timeout';
        case 'expected-error':
        case 'expectederror':
            return 'expected-error';
        case 'reason':
            return 'reason';
        default:
            return undefined;
    }
}

function metadataKeyIsSupportedForDirective(
    key: TestDirectiveMetadataKey,
    kind: TestDirectiveKind,
): boolean {
    return key !== 'reason' || kind === 'skip' || kind === 'xfail';
}

function isInsideMetadataQuote(text: string): boolean {
    let quote: '"' | "'" | undefined;
    for (const char of text) {
        if (char !== '"' && char !== "'") {
            continue;
        }
        if (!quote) {
            quote = char;
        } else if (quote === char) {
            quote = undefined;
        }
    }
    return quote !== undefined;
}
