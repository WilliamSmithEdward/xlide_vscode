import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { parseModule, parseModuleFreshForTests } from '../src/analyzer/parser/parseModule';
import { moduleHasConditionalDirectives } from '../src/analyzer/conditional/conditionalCompilation';

afterEach(() => vi.restoreAllMocks());

describe('unchanged module snapshot facts', () => {
    it.each(['', 'DefLng A-Z\n'])('scans DefType facts once across symbol projections for %s', header => {
        const source = header + "' unchanged source facts probe\n" + Array.from({ length: 1200 }, (_, i) =>
            'Sub SourceFactPadding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('');
        const parsedModule = parseModule(source);
        const native = String.prototype.matchAll;
        let scans = 0;
        vi.spyOn(String.prototype, 'matchAll').mockImplementation(function(this: string, regex: RegExp) {
            if (regex.source.includes('Def(Bool|Byte|Int|')) { scans++; }
            return native.call(this, regex);
        });
        for (const moduleName of ['CandidateA', 'CandidateB', 'CandidateC']) {
            const result = buildModuleSymbols(moduleName, 'class', source, { parsedModule });
            expect(result.defTypes?.get('a')).toBe(header ? 'Long' : undefined);
        }
        expect(scans).toBe(1);
    });

    it('uses the parser directive-free fact without revisiting unchanged procedure bodies', () => {
        const source = "' directive free facts probe\n" + Array.from({ length: 1200 }, (_, i) =>
            'Sub DirectiveFactPadding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('');
        const module = parseModule(source);
        let reads = 0;
        const restore: (() => void)[] = [];
        try {
            for (const member of module.members) {
                if (member.kind !== 'Procedure') { continue; }
                const descriptor = Object.getOwnPropertyDescriptor(member, 'body')!;
                const body = member.body;
                Object.defineProperty(member, 'body', { configurable: true, get() { reads++; return body; } });
                restore.push(() => Object.defineProperty(member, 'body', descriptor));
            }
            expect(moduleHasConditionalDirectives(module)).toBe(false);
            expect(moduleHasConditionalDirectives(module)).toBe(false);
            expect(reads).toBe(0);
        } finally {
            restore.forEach(run => run());
        }
    });
});

describe('snapshot fact validity and ownership', () => {
    it('recomputes DefType facts when text changes and keeps returned maps separate', () => {
        const source = "DefInt A-M\nDefStr N-Z\n' isolated DefType maps\nSub Probe()\nEnd Sub\n";
        const first = buildModuleSymbols('First', 'standard', source);
        (first.defTypes as Map<string, string>).set('a', 'Mutated');
        const second = buildModuleSymbols('Second', 'class', source);
        expect(second.defTypes?.get('a')).toBe('Integer');
        expect(second.defTypes?.get('n')).toBe('String');
        expect(second.defTypes).not.toBe(first.defTypes);
        const changed = buildModuleSymbols('Second', 'class', source.replace('DefInt', 'DefLng'));
        expect(changed.defTypes?.get('a')).toBe('Long');
        expect(changed.defTypes?.get('n')).toBe('String');
        expect(buildModuleSymbols('Again', 'standard', source).defTypes?.get('a')).toBe('Integer');
    });

    it('bounds retained DefType snapshots and rescans an evicted source', () => {
        const source = "DefByte A-Z\n' bounded DefType retention\nSub Probe()\nEnd Sub\n";
        buildModuleSymbols('First', 'standard', source);
        for (let i = 0; i < 8; i++) { buildModuleSymbols('Padding', 'standard', source + "' snapshot " + i); }
        const native = String.prototype.matchAll;
        let scans = 0;
        vi.spyOn(String.prototype, 'matchAll').mockImplementation(function(this: string, regex: RegExp) {
            if (regex.source.includes('Def(Bool|Byte|Int|')) { scans++; }
            return native.call(this, regex);
        });
        expect(buildModuleSymbols('Restored', 'standard', source).defTypes?.get('z')).toBe('Byte');
        expect(scans).toBe(1);
    });

    it.each([
        'Sub P()\nDebug.Print 1\nEnd Sub',
        '#Const ENABLED = True\nSub P()\nEnd Sub',
        'Sub P()\n#If VBA7 Then\nDebug.Print 1\n#End If\nEnd Sub',
        'Enum Choice\n#If VBA7 Then\nOne\n#End If\nEnd Enum',
        'Type Record\n#If VBA7 Then\nvalue As Long\n#End If\nEnd Type',
        'Sub P()\n#If VBA7 Then\nDebug.Print 1',
        'Sub P()\nDebug.Print "#If VBA7 Then"\nEnd Sub',
        "Sub P()\n' #Const TEST = True\nDebug.Print #1/1/2000#\nEnd Sub",
        '#Unknown BAD\nSub P()\nEnd Sub',
    ])('matches an independent AST traversal for %s', source => {
        expect(moduleHasConditionalDirectives(parseModule(source)))
            .toBe(moduleHasConditionalDirectives(parseModuleFreshForTests(source)));
    });

    it('preserves directive presence while adding and removing a directive in a large body', () => {
        const prefix = Array.from({ length: 500 }, (_, i) =>
            'Sub IncrementalFactPadding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('');
        const source = prefix + 'Sub SnapshotProbe()\nDebug.Print 314159\nEnd Sub\n';
        const withDirective = source.replace('Debug.Print 314159', '#Const ADDED = True');
        const changedFree = source.replace('Debug.Print 314159', 'Debug.Print 271828');
        for (const [text, expected] of [[source, false], [withDirective, true], [changedFree, false], [source, false]] as const) {
            const module = parseModule(text);
            expect(moduleHasConditionalDirectives(module)).toBe(expected);
            expect(moduleHasConditionalDirectives(module)).toBe(moduleHasConditionalDirectives(parseModuleFreshForTests(text)));
        }
    });

    it('traverses caller-supplied trees even when based on a known directive-free module', () => {
        const cached = parseModule("Sub CallerSuppliedProbe()\nEnd Sub\n");
        const members = parseModuleFreshForTests('#Const SUPPLIED = True\n').members;
        expect(moduleHasConditionalDirectives({ ...cached, members })).toBe(true);
        members.splice(0);
        expect(moduleHasConditionalDirectives({ ...cached, members })).toBe(false);
        expect(moduleHasConditionalDirectives(cached)).toBe(false);
    });
});
