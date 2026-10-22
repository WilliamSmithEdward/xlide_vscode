/* Matching runs only in a terminable worker, never on the extension host. */
function searchModule(text, query, limit) {
    if (!query.pattern || query.pattern.length > 1024) { throw new Error('Enter a pattern of 1–1024 characters.'); }
    const pattern = query.regex ? query.pattern : query.pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(pattern, query.matchCase ? 'gmu' : 'gimu');
    const word = ch => ch !== undefined && /[\p{L}\p{M}\p{N}_]/u.test(ch);
    const before = index => {
        if (!index) { return undefined; }
        const last = text.charCodeAt(index - 1);
        return text.slice(index - (last >= 0xDC00 && last <= 0xDFFF ? 2 : 1), index);
    };
    const after = index => index < text.length ? String.fromCodePoint(text.codePointAt(index)) : undefined;
    const matches = [];
    let captureLength = 0;
    let match;
    while ((match = re.exec(text)) !== null) {
        const start = match.index, end = start + match[0].length;
        if (!query.wholeWord || (!word(before(start)) && !word(after(end)))) {
            if (matches.length >= limit) { return { matches, limitHit: true }; }
            captureLength += Array.from(match, value => value?.length ?? 0).reduce((a, b) => a + b, 0);
            if (captureLength > 2_000_000) { throw new Error('Search captures are too large. Narrow the scope.'); }
            matches.push({ start, end, captures: Array.from(match, value => value ?? null), named: match.groups ?? {} });
        }
        if (!match[0].length) {
            if (re.lastIndex >= text.length) { break; }
            re.lastIndex += text.codePointAt(re.lastIndex) > 0xFFFF ? 2 : 1;
        }
    }
    return { matches, limitHit: false };
}

if (typeof self !== 'undefined') {
    self.onmessage = event => {
        if (event.origin !== '') { return; }
        const { data } = event;
        try { self.postMessage({ id: data.id, result: searchModule(data.text, data.query, data.limit) }); }
        catch (error) { self.postMessage({ id: data.id, error: error.message }); }
    };
}
if (typeof module !== 'undefined') { module.exports = { searchModule }; }
