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
}

const TEST_DIRECTIVE_COMPLETIONS: readonly Omit<VbaTestDirectiveCompletion, 'range'>[] = [
    {
        label: '@xlide-test',
        insertText: '@xlide-test',
        detail: 'XLIDE VBA test',
        documentation: 'Marks the following no-argument standard-module Sub as an XLIDE test.',
        sortText: '0:@xlide-test',
    },
    {
        label: '@xlide-test-skip',
        insertText: '@xlide-test-skip reason="$1"',
        detail: 'XLIDE skipped VBA test',
        documentation: 'Marks the following XLIDE test as skipped and records the reason.',
        sortText: '1:@xlide-test-skip',
    },
    {
        label: '@xlide-test-xfail',
        insertText: '@xlide-test-xfail reason="$1"',
        detail: 'XLIDE expected-failure VBA test',
        documentation: 'Marks the following XLIDE test as expected to fail and records the reason.',
        sortText: '2:@xlide-test-xfail',
    },
];

const DIRECTIVE_TOKEN_RE = /^@?[A-Za-z0-9-]*$/;
const DIRECTIVE_SUFFIX_RE = /[A-Za-z0-9-]/;

export function resolveVbaTestDirectiveCompletions(
    line: string,
    character: number,
): VbaTestDirectiveCompletion[] {
    const cursor = Math.max(0, Math.min(character, line.length));
    const before = line.slice(0, cursor);
    const comment = /^\s*'(?!'')\s*/.exec(before);
    if (!comment) {
        return [];
    }

    const start = comment[0].length;
    const typed = before.slice(start);
    if (!DIRECTIVE_TOKEN_RE.test(typed)) {
        return [];
    }

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
        ...completion,
        range: { start, end },
    }));
}
