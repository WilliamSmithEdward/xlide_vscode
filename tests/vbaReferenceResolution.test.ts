import { describe, expect, it } from 'vitest';
import { collectSymbolReferences } from '../src/vbaReferenceResolution';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';
import { lineStartOffsets } from '../src/vbaSourceScan';

interface Mod {
    moduleName: string;
    source: string;
    type?: string;
    documentType?: undefined;
}

function setup(mods: Mod[]) {
    const project = buildVbaProjectIndex(
        mods.map((m) => ({ moduleName: m.moduleName, source: m.source, type: m.type })),
    );
    const modules = mods.map((m) => ({
        moduleName: m.moduleName,
        source: m.source,
        type: m.type,
        documentType: m.documentType,
    }));
    const byModule = new Map(modules.map((m) => [m.moduleName.toLowerCase(), m]));
    return { project, modules, byModule };
}

/**
 * Resolves references to the symbol whose name `word` starts at `atOffset` in
 * `currentModule`, then applies the rename to `newName` and returns each
 * module's rewritten source. This exercises the full provider resolution path.
 */
function applyRename(
    mods: Mod[],
    currentModule: string,
    word: string,
    atOffset: number,
    newName: string,
): Record<string, string> {
    const { project, modules, byModule } = setup(mods);
    const current = byModule.get(currentModule.toLowerCase())!;
    const result = collectSymbolReferences(
        byModule,
        project,
        modules,
        current.source,
        currentModule,
        current,
        word,
        atOffset + word.length,
        atOffset,
        true,
    );
    expect(result.hasSymbol).toBe(true);
    const out: Record<string, string> = {};
    for (const m of modules) {
        const edits = result.references
            .filter((r) => r.moduleName.toLowerCase() === m.moduleName.toLowerCase())
            .map((r) => ({ offset: (lineStartOffsets(m.source)[r.line] ?? 0) + r.column, length: r.length }))
            .sort((a, b) => b.offset - a.offset);
        let s = m.source;
        for (const e of edits) {
            s = s.slice(0, e.offset) + newName + s.slice(e.offset + e.length);
        }
        out[m.moduleName] = s;
    }
    return out;
}

describe('collectSymbolReferences - standard-module procedure rename', () => {
    const MODULE1 =
        'Sub Greet(a As String)\n' +
        'End Sub\n' +
        'Sub Local()\n' +
        '    Greet "x"\n' +
        '    Greet("x")\n' +
        '    Call Greet("x")\n' +
        'End Sub\n';
    const MODULE2 =
        'Sub Cross()\n' +
        '    Greet "y"\n' +
        '    Module1.Greet "y"\n' +
        'End Sub\n';

    const mods: Mod[] = [
        { moduleName: 'Module1', source: MODULE1 },
        { moduleName: 'Module2', source: MODULE2 },
    ];

    const EXPECT_M1 =
        'Sub Hello(a As String)\n' +
        'End Sub\n' +
        'Sub Local()\n' +
        '    Hello "x"\n' +
        '    Hello("x")\n' +
        '    Call Hello("x")\n' +
        'End Sub\n';
    const EXPECT_M2 =
        'Sub Cross()\n' +
        '    Hello "y"\n' +
        '    Module1.Hello "y"\n' +
        'End Sub\n';

    it('renames the declaration plus every bare call form (same- and cross-module) and the qualified call, when invoked from the declaration', () => {
        const out = applyRename(mods, 'Module1', 'Greet', MODULE1.indexOf('Greet'), 'Hello');
        expect(out.Module1).toBe(EXPECT_M1);
        expect(out.Module2).toBe(EXPECT_M2);
    });

    it('produces the identical result when invoked from a bare call site', () => {
        const out = applyRename(mods, 'Module2', 'Greet', MODULE2.indexOf('Greet'), 'Hello');
        expect(out.Module1).toBe(EXPECT_M1);
        expect(out.Module2).toBe(EXPECT_M2);
    });

    it('produces the identical result when invoked from the module-qualified reference', () => {
        const at = MODULE2.indexOf('Greet', MODULE2.indexOf('Module1.'));
        const out = applyRename(mods, 'Module2', 'Greet', at, 'Hello');
        expect(out.Module1).toBe(EXPECT_M1);
        expect(out.Module2).toBe(EXPECT_M2);
    });

    it('does not over-match a same-named member on an unrelated class receiver', () => {
        const widget = 'Public Sub Greet()\nEnd Sub\n';
        const user = 'Sub U()\n    Dim w As Widget\n    w.Greet\nEnd Sub\n';
        const out = applyRename(
            [
                { moduleName: 'Module1', source: MODULE1 },
                { moduleName: 'Widget', source: widget, type: 'class' },
                { moduleName: 'Module3', source: user },
            ],
            'Module1',
            'Greet',
            MODULE1.indexOf('Greet'),
            'Hello',
        );
        expect(out.Module1).toBe(EXPECT_M1);
        // The class's own member and the w.Greet call are a different symbol.
        expect(out.Widget).toBe(widget);
        expect(out.Module3).toBe(user);
    });
});

describe('collectSymbolReferences - class member rename', () => {
    const PERSON =
        'Public Sub Save()\n' +
        '    Save\n' +
        '    Me.Save\n' +
        'End Sub\n' +
        'Public Sub DoIt()\n' +
        '    Save\n' +
        'End Sub\n';
    const CALLER =
        'Sub U()\n' +
        '    Dim p As Person\n' +
        '    p.Save\n' +
        'End Sub\n';
    // An unrelated standard-module Sub that shares the name must be untouched.
    const OTHER = 'Sub Save()\nEnd Sub\nSub V()\n    Save\nEnd Sub\n';

    const mods: Mod[] = [
        { moduleName: 'Person', source: PERSON, type: 'class' },
        { moduleName: 'Caller', source: CALLER },
        { moduleName: 'Other', source: OTHER },
    ];

    it('renames the declaration, bare in-class calls, Me.member, and obj.member, but not an unrelated same-named standard Sub', () => {
        const out = applyRename(mods, 'Person', 'Save', PERSON.indexOf('Save'), 'Persist');
        expect(out.Person).toBe(
            'Public Sub Persist()\n' +
            '    Persist\n' +
            '    Me.Persist\n' +
            'End Sub\n' +
            'Public Sub DoIt()\n' +
            '    Persist\n' +
            'End Sub\n',
        );
        expect(out.Caller).toBe('Sub U()\n    Dim p As Person\n    p.Persist\nEnd Sub\n');
        expect(out.Other).toBe(OTHER);
    });
});

describe('collectSymbolReferences - property accessors', () => {
    const PERSON =
        'Private mName As String\n' +
        'Public Property Get Name() As String\n' +
        '    Name = mName\n' +
        'End Property\n' +
        'Public Property Let Name(ByVal v As String)\n' +
        '    mName = v\n' +
        'End Property\n';
    const CALLER =
        'Sub U()\n' +
        '    Dim p As Person\n' +
        '    p.Name = "a"\n' +
        '    Debug.Print p.Name\n' +
        'End Sub\n';

    it('renames all accessors and qualified uses', () => {
        const out = applyRename(
            [
                { moduleName: 'Person', source: PERSON, type: 'class' },
                { moduleName: 'Caller', source: CALLER },
            ],
            'Person',
            'Name',
            PERSON.indexOf('Name'),
            'FullName',
        );
        expect(out.Person).toContain('Property Get FullName()');
        expect(out.Person).toContain('Property Let FullName(');
        expect(out.Person).toContain('FullName = mName');
        expect(out.Caller).toBe('Sub U()\n    Dim p As Person\n    p.FullName = "a"\n    Debug.Print p.FullName\nEnd Sub\n');
    });
});

describe('collectSymbolReferences - Declare statement', () => {
    it('renames a Declare alias and its bare calls', () => {
        const M1 =
            'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
            'Sub T()\n' +
            '    Sleep 100\n' +
            '    Sleep 200\n' +
            'End Sub\n';
        const out = applyRename([{ moduleName: 'Module1', source: M1 }], 'Module1', 'Sleep', M1.indexOf('Sleep'), 'Delay');
        expect(out.Module1).toBe(
            'Declare PtrSafe Sub Delay Lib "kernel32" (ByVal ms As Long)\n' +
            'Sub T()\n' +
            '    Delay 100\n' +
            '    Delay 200\n' +
            'End Sub\n',
        );
    });
});

describe('collectSymbolReferences - find references excludes the declaration', () => {
    it('omits the declaration span when includeDeclaration is false', () => {
        const M1 = 'Sub Greet()\nEnd Sub\nSub Caller()\n    Greet\nEnd Sub\n';
        const { project, modules, byModule } = setup([{ moduleName: 'Module1', source: M1 }]);
        const declAt = M1.indexOf('Greet');
        const callAt = M1.indexOf('Greet', M1.indexOf('Caller'));
        const result = collectSymbolReferences(
            byModule, project, modules, M1, 'Module1', byModule.get('module1'),
            'Greet', declAt + 5, declAt, false,
        );
        const offsets = result.references
            .map((r) => (lineStartOffsets(M1)[r.line] ?? 0) + r.column)
            .sort((a, b) => a - b);
        expect(offsets).toEqual([callAt]);
    });
});

describe('collectSymbolReferences - ambiguous same-named public procedures', () => {
    it('renames the targeted module-qualified reference but leaves an ambiguous bare call alone', () => {
        const M1 = 'Public Sub Greet()\nEnd Sub\n';
        const M2 = 'Public Sub Greet()\nEnd Sub\n';
        const M3 = 'Sub U()\n    Module1.Greet\n    Greet\nEnd Sub\n';
        const out = applyRename(
            [
                { moduleName: 'Module1', source: M1 },
                { moduleName: 'Module2', source: M2 },
                { moduleName: 'Module3', source: M3 },
            ],
            'Module1',
            'Greet',
            M1.indexOf('Greet'),
            'Hello',
        );
        expect(out.Module1).toBe('Public Sub Hello()\nEnd Sub\n');
        expect(out.Module2).toBe(M2);
        // Module1.Greet -> renamed; the ambiguous bare Greet is left untouched.
        expect(out.Module3).toBe('Sub U()\n    Module1.Hello\n    Greet\nEnd Sub\n');
    });
});

describe('collectSymbolReferences - Property Set accessor', () => {
    it('renames Get, Let, and Set accessors together', () => {
        const PERSON =
            'Public Property Get Item() As Object\nEnd Property\n' +
            'Public Property Let Item(ByVal v As Variant)\nEnd Property\n' +
            'Public Property Set Item(ByVal v As Object)\nEnd Property\n';
        const out = applyRename(
            [{ moduleName: 'Bag', source: PERSON, type: 'class' }],
            'Bag',
            'Item',
            PERSON.indexOf('Item'),
            'Value',
        );
        expect(out.Bag).toBe(
            'Public Property Get Value() As Object\nEnd Property\n' +
            'Public Property Let Value(ByVal v As Variant)\nEnd Property\n' +
            'Public Property Set Value(ByVal v As Object)\nEnd Property\n',
        );
    });
});

describe('collectSymbolReferences - module-level variable', () => {
    it('renames bare same-module uses and module-qualified cross-module uses', () => {
        const M1 = 'Public Counter As Long\nSub Bump()\n    Counter = Counter + 1\nEnd Sub\n';
        const M2 = 'Sub R()\n    Module1.Counter = 5\nEnd Sub\n';
        const out = applyRename(
            [
                { moduleName: 'Module1', source: M1 },
                { moduleName: 'Module2', source: M2 },
            ],
            'Module1',
            'Counter',
            M1.indexOf('Counter'),
            'Total',
        );
        expect(out.Module1).toBe('Public Total As Long\nSub Bump()\n    Total = Total + 1\nEnd Sub\n');
        expect(out.Module2).toBe('Sub R()\n    Module1.Total = 5\nEnd Sub\n');
    });
});

describe('collectSymbolReferences - local variable scope', () => {
    it('renames only the local in its own procedure, not a same-named local elsewhere', () => {
        const M1 =
            'Sub T()\n    Dim x As Long\n    x = 1\n    Debug.Print x\nEnd Sub\n' +
            'Sub U()\n    Dim x As Long\n    x = 2\nEnd Sub\n';
        const out = applyRename([{ moduleName: 'Module1', source: M1 }], 'Module1', 'x', M1.indexOf('Dim x') + 4, 'y');
        expect(out.Module1).toBe(
            'Sub T()\n    Dim y As Long\n    y = 1\n    Debug.Print y\nEnd Sub\n' +
            'Sub U()\n    Dim x As Long\n    x = 2\nEnd Sub\n',
        );
    });
});

describe('collectSymbolReferences - user-defined Type field', () => {
    it('renames a Type field declaration and its member-access uses', () => {
        const M1 =
            'Public Type Pt\n    X As Long\n    Y As Long\nEnd Type\n' +
            'Sub T()\n    Dim p As Pt\n    p.X = 1\n    Debug.Print p.X\nEnd Sub\n';
        const out = applyRename([{ moduleName: 'Module1', source: M1 }], 'Module1', 'X', M1.indexOf('X As Long'), 'Col');
        expect(out.Module1).toBe(
            'Public Type Pt\n    Col As Long\n    Y As Long\nEnd Type\n' +
            'Sub T()\n    Dim p As Pt\n    p.Col = 1\n    Debug.Print p.Col\nEnd Sub\n',
        );
    });
});

describe('collectSymbolReferences - Enum member', () => {
    // Enum-TYPE-qualified access used to be left untouched: the analyzer had no
    // EnumType.Member resolver, so `Color.Red` could only have been matched
    // textually, which would over-match an unrelated OtherEnum.Red. Enums are
    // member surfaces now (issue #11), so the qualified use resolves properly
    // and renames with the declaration - leaving it behind pointed the module
    // at a member that no longer exists.
    it('renames an Enum member declaration plus its bare and qualified uses', () => {
        const M1 =
            'Public Enum Color\n    Red\n    Green\nEnd Enum\n' +
            'Sub T()\n    Dim c As Long\n    c = Red\n    c = Color.Red\nEnd Sub\n';
        const out = applyRename([{ moduleName: 'Module1', source: M1 }], 'Module1', 'Red', M1.indexOf('Red'), 'Crimson');
        expect(out.Module1).toBe(
            'Public Enum Color\n    Crimson\n    Green\nEnd Enum\n' +
            'Sub T()\n    Dim c As Long\n    c = Crimson\n    c = Color.Crimson\nEnd Sub\n',
        );
    });
});

describe('collectSymbolReferences - non-symbol', () => {
    it('reports hasSymbol=false for an unresolved identifier', () => {
        const { project, modules, byModule } = setup([{ moduleName: 'Module1', source: 'Sub T()\n    Foo Bar\nEnd Sub\n' }]);
        const src = byModule.get('module1')!.source;
        const at = src.indexOf('Bar');
        const result = collectSymbolReferences(
            byModule, project, modules, src, 'Module1', byModule.get('module1'),
            'Bar', at + 3, at, true,
        );
        expect(result.hasSymbol).toBe(false);
        expect(result.references).toEqual([]);
    });
});
