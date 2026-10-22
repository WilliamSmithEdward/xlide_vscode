import { lineStartOffsets } from './vbaSourceScan';

export interface ModuleSearchQuery {
    pattern: string;
    regex: boolean;
    matchCase: boolean;
    wholeWord: boolean;
}

export interface ModuleSearchMatch {
    start: number;
    end: number;
    captures: Array<string | null>;
    named: Record<string, string | undefined>;
}

/** Validate worker messages before retaining positions or preparing edits. */
export function validateModuleSearchMatches(source: string, value: unknown, limit: number): ModuleSearchMatch[] {
    if (!Array.isArray(value) || value.length > limit) { throw new Error('Invalid search result count.'); }
    let end = 0;
    let previousStart = -1;
    for (const match of value) {
        if (!match || !Number.isInteger(match.start) || !Number.isInteger(match.end)
            || match.start < end || match.start <= previousStart || match.end < match.start || match.end > source.length
            || !Array.isArray(match.captures) || match.captures.length > 1024
            || !match.captures.every((c: unknown) => c === null || typeof c === 'string')
            || match.captures[0] !== source.slice(match.start, match.end)
            || !match.named || typeof match.named !== 'object' || Array.isArray(match.named)
            || Object.keys(match.named).length > 1024
            || !Object.values(match.named).every(c => c === undefined || typeof c === 'string')) {
            throw new Error('Invalid search result range.');
        }
        end = match.end;
        previousStart = match.start;
    }
    return value;
}

/** JavaScript replacement substitutions, using captures from the original input. */
export function moduleSearchReplacement(source: string, match: ModuleSearchMatch, replacement: string, regex: boolean, limit = 20_000_000): string {
    if (replacement.length > limit) { throw new Error('Replacement size limit reached. Narrow the scope.'); }
    if (!regex) { return replacement; }
    const substitute = (literal: string, token: string): string => {
        if (token === '$') { return '$'; }
        if (token === '&') { return match.captures[0] ?? ''; }
        if (token === '`') { return source.slice(0, match.start); }
        if (token === "'") { return source.slice(match.end); }
        if (token.startsWith('<')) {
            if (!Object.keys(match.named).length) { return literal; }
            const name = token.slice(1, -1);
            return Object.prototype.hasOwnProperty.call(match.named, name) ? match.named[name] ?? '' : '';
        }
        let index = Number(token);
        if (index > 0 && index < match.captures.length) { return match.captures[index] ?? ''; }
        if (token.length === 2) {
            index = Number(token[0]);
            if (index > 0 && index < match.captures.length) { return (match.captures[index] ?? '') + token[1]; }
        }
        return literal;
    };
    let length = replacement.length;
    return replacement.replace(/\$(\$|&|`|'|<[^>]*>|\d{1,2})/g, (literal, token: string) => {
        const value = substitute(literal, token);
        length += value.length - literal.length;
        if (length > limit) { throw new Error('Replacement size limit reached. Narrow the scope.'); }
        return value;
    });
}

/** Bounded context around the entire edit, including multiline and zero-width matches. */
export function moduleSearchChangePreview(source: string, match: ModuleSearchMatch, replacement: string) {
    const prefixWindow = source.slice(Math.max(0, match.start - 100), match.start);
    const prefix = prefixWindow.slice(prefixWindow.lastIndexOf('\n') + 1);
    const suffixWindow = source.slice(match.end, match.end + 120);
    const nextLine = suffixWindow.indexOf('\n');
    const suffix = (nextLine < 0 ? suffixWindow : suffixWindow.slice(0, nextLine)).replace(/\r$/, '');
    const prefixStart = match.start - prefix.length;
    const suffixEnd = match.end + (nextLine < 0 ? suffixWindow.length : nextLine);
    const clip = (text: string) => ({ text: text.slice(0, 500), truncated: text.length > 500 });
    return {
        prefix,
        suffix,
        before: clip(source.slice(match.start, match.end)),
        after: clip(replacement),
        contextTruncated: prefixStart > 0 && source[prefixStart - 1] !== '\n' || suffixEnd < source.length && source[suffixEnd] !== '\n',
    };
}

export function moduleSearchPreviewer(source: string): (match: ModuleSearchMatch) => {
    line: number; text: string; highlightStart: number; highlightEnd: number;
} {
    const starts = lineStartOffsets(source);
    return match => {
        let low = 0, high = starts.length;
        while (low + 1 < high) {
            const mid = (low + high) >>> 1;
            if (starts[mid] <= match.start) { low = mid; } else { high = mid; }
        }
        const start = Math.max(starts[low], match.start - 100);
        const end = Math.min(starts[low + 1] ?? source.length, match.start + 120);
        const text = source.slice(start, end).replace(/[\r\n]+$/, '');
        return { line: low + 1, text, highlightStart: match.start - start,
            highlightEnd: Math.min(text.length, match.end - start) };
    };
}
