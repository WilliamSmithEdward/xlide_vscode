import { describe, expect, it } from 'vitest';
import { privateMemberOwnerAt, projectClassMemberAt, type MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const source = 'Sub P()\nMe.Hit\nEnd Sub';
const offset = source.indexOf('Me.Hit') + 'Me.Hit'.length;
const context = (owner: VbaProjectClassMembers): MemberCompletionContext => ({
    meProjectType: owner.name, projectClassMembers: [owner],
});

describe('project member lookup snapshots', () => {
    it('keeps case-insensitive first-match behavior and refreshed member arrays', () => {
        const first = { name: 'hIT', kind: 'method' as const, moduleName: 'C', returns: 'Long' };
        const duplicate = { ...first, name: 'Hit', returns: 'String' };
        const owner: VbaProjectClassMembers = { name: 'C', moduleName: 'C', kind: 'class', members: [first, duplicate] };
        expect(projectClassMemberAt(source, offset, 'HIT', context(owner))).toBe(first);
        owner.members = [duplicate];
        expect(projectClassMemberAt(source, offset, 'hit', context(owner))).toBe(duplicate);
        expect(projectClassMemberAt(source, offset, 'missing', context(owner))).toBeUndefined();
        expect(projectClassMemberAt(source, offset, 'hit', context({ ...owner, kind: 'document' }))).toBeUndefined();
    });

    it('refreshes private names and lets a public surface of the same name win', () => {
        const owner: VbaProjectClassMembers = { name: 'C', moduleName: 'C', kind: 'class', members: [], privateMembers: ['hIT'] };
        expect(privateMemberOwnerAt(source, offset, 'HIT', context(owner))).toBe('C');
        owner.privateMembers = ['Other'];
        expect(privateMemberOwnerAt(source, offset, 'hit', context(owner))).toBeUndefined();
        owner.privateMembers = ['Hit'];
        owner.members = [{ name: 'Hit', kind: 'method', moduleName: 'C' }];
        expect(privateMemberOwnerAt(source, offset, 'hit', context(owner))).toBeUndefined();
    });

    it('reads a large public member list once across repeated misses', () => {
        let reads = 0;
        const members = Array.from({ length: 1000 }, (_, i) => ({
            get name() { reads++; return `Member${i}`; }, kind: 'method' as const, moduleName: 'C',
        }));
        const owner: VbaProjectClassMembers = { name: 'C', moduleName: 'C', kind: 'class', members };
        const ctx = context(owner);
        for (let i = 0; i < 100; i++) expect(projectClassMemberAt(source, offset, 'missing', ctx)).toBeUndefined();
        expect(reads).toBe(1000);
    });

    it('reads a large private name list once across repeated misses', () => {
        let reads = 0;
        const privateMembers = Array.from({ length: 1000 }, (_, i) => `Member${i}`);
        for (let i = 0; i < privateMembers.length; i++) {
            Object.defineProperty(privateMembers, i, { get() { reads++; return `Member${i}`; } });
        }
        const owner: VbaProjectClassMembers = { name: 'C', moduleName: 'C', kind: 'class', members: [], privateMembers };
        const ctx = context(owner);
        for (let i = 0; i < 100; i++) expect(privateMemberOwnerAt(source, offset, 'missing', ctx)).toBeUndefined();
        expect(reads).toBe(1000);
    });
});
