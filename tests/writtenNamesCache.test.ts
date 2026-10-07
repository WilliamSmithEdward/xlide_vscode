import { describe, expect, it } from 'vitest';
import { writtenNamesIn } from '../src/analyzer/diagnostics/moduleState';

describe('bounded source write facts', () => {
    it('reuses equal source content and refreshes edited assignment targets', () => {
        const source = 'Sub WriteCacheProbe()\nfirst = 1\nEnd Sub';
        const facts = writtenNamesIn(source);
        expect(facts.has('first')).toBe(true);
        expect(writtenNamesIn(JSON.parse(JSON.stringify(source)))).toBe(facts);
        const edited = source.replace('first = 1', 'second = 1');
        expect(writtenNamesIn(edited).has('first')).toBe(false);
        expect(writtenNamesIn(edited).has('second')).toBe(true);
        expect(writtenNamesIn(source)).toBe(facts);
    });

    it('bounds retained sources while keeping a recently read one', () => {
        const source = 'Sub RecentWriteCacheProbe()\nkept = 1\nEnd Sub';
        const facts = writtenNamesIn(source);
        for (let i = 0; i < 8; i++) {
            writtenNamesIn(`Sub OtherWriteCacheProbe${i}()\nother${i} = 1\nEnd Sub`);
            expect(writtenNamesIn(source)).toBe(facts);
        }
        for (let i = 0; i < 8; i++) {
            writtenNamesIn(`Sub EvictWriteCacheProbe${i}()\nevict${i} = 1\nEnd Sub`);
        }
        const rebuilt = writtenNamesIn(source);
        expect(rebuilt).not.toBe(facts);
        expect([...rebuilt]).toEqual([...facts]);
    });
});
