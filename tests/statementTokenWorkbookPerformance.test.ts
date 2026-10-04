import { expect, it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { statementTokensCached } from '../src/analyzer/lexer/tokenHelpers';
function spans(source: string) {
    const result: { start: number; end: number }[] = [];
    for (const member of parseModule(source).members) {
        if (member.kind === 'Procedure') forEachStatementWithHeaders(source, member.body, statement => result.push(statement.span));
    }
    return result;
}
it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('measures fresh edit statement token construction with cold parity in the actual class', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const originalSpans = spans(original);
    const after: number[] = [];
    for (let i = 0; i < 21; i++) {
        originalSpans.map(span => statementTokensCached(original, span));
        const source = original + "\r\nPublic Sub XlideTokenProbe()\r\n Debug.Print " + i + '\r\nEnd Sub\r\n';
        const next = spans(source);
        tokenizeCached(source);
        const newRun = () => { const start = performance.now(); const result = next.map(span => statementTokensCached(source, span)); after.push(performance.now() - start); return result; };
        const current = newRun();
        // Evict the two-entry statement cache, then rebuild the same spans.
        for (let j = 0; j < 3; j++) statementTokensCached("' eviction " + j, { start: 0, end: 1 });
        expect(current).toEqual(next.map(span => statementTokensCached(source, span)));
    }
    after.sort((a,b)=>a-b);
    process.stdout.write('Actual statement token construction: ' + JSON.stringify({ statements: originalSpans.length, afterMedian: after[10], afterMax: after[20] }) + '\n');
}, 60000);
