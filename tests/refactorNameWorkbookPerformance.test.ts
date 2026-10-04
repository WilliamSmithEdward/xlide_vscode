import { expect, it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { nameAt } from '../src/analyzer/refactor/shared';
function precedingLookup(source: string, offset: number): string | undefined {
    const before = /[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.exec(source.slice(0, offset));
    const after = /^[\p{L}\p{M}\p{N}_]*/u.exec(source.slice(offset));
    const name = `${before?.[0] ?? ''}${after?.[0] ?? ''}`;
    return /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.test(name) ? name : undefined;
}
it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('compares line-bounded identifier lookup in the actual class', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const source = original + '\r\nPublic Sub XlideNameProbe()\r\n Debug.Print LatencyValue\r\nEnd Sub\r\n';
    const offset = source.lastIndexOf('LatencyValue') + 4;
    const before: number[] = [], after: number[] = [];
    for (let i = 0; i < 21; i++) {
        const oldRun = () => { const start = performance.now(); for (let j = 0; j < 10; j++) precedingLookup(source, offset); before.push((performance.now() - start) / 10); };
        const newRun = () => { const start = performance.now(); for (let j = 0; j < 10; j++) nameAt(source, offset); after.push((performance.now() - start) / 10); };
        if (i % 2) { oldRun(); newRun(); } else { newRun(); oldRun(); }
        expect(nameAt(source, offset)).toBe(precedingLookup(source, offset));
    }
    before.sort((a,b)=>a-b); after.sort((a,b)=>a-b);
    process.stdout.write('Actual refactor name lookup (ms): ' + JSON.stringify({ beforeMedian: before[10], afterMedian: after[10], beforeMax: before[20], afterMax: after[20] }) + '\n');
});
