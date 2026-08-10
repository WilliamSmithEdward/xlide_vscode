// The `Interface_Member` prefix on an implementing class.
//
// `Private Sub IShape_Draw()` is how VBA binds a class's implementation to the
// interface it declares with `Implements IShape`. The prefix is a contract, not
// a naming convention: rename the interface and leave the prefix behind and the
// class stops implementing anything, which the compiler notices and the
// developer did not ask for.
//
// Kept free of any `vscode` import so the scan can be unit-tested directly.

/** Where an interface prefix appears in a module's source. */
export interface InterfacePrefixHit {
    /** 0-based physical line. */
    line: number;
    /** 0-based column of the interface name itself, not the whole procedure name. */
    column: number;
    /** Length of the interface name at that position. */
    length: number;
}

/**
 * Finds `<interfaceName>_` prefixes on procedure declarations in `source`.
 *
 * Only declarations count. A variable, a string or a comment that happens to
 * start with the same word is not part of the contract and is left alone, and
 * the caller must only pass modules that actually implement the interface -
 * `IShapeLookalike` in an unrelated module means nothing.
 */
export function interfacePrefixHits(source: string, interfaceName: string): InterfacePrefixHit[] {
    const hits: InterfacePrefixHit[] = [];
    const escaped = interfaceName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // The prefix must be followed by `_` and a further identifier character, so
    // a procedure named exactly after the interface is not mistaken for one.
    const declaration = new RegExp(
        String.raw`^(\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?`
        + String.raw`(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+)(${escaped})(_[\p{L}\p{M}\p{N}_])`,
        'iu',
    );
    const lines = source.split(/\r?\n/);
    for (let line = 0; line < lines.length; line += 1) {
        const match = declaration.exec(lines[line]);
        if (!match) {
            continue;
        }
        hits.push({ line, column: match[1].length, length: match[2].length });
    }
    return hits;
}

/** Applies the prefix rename to a module's source, for callers holding text. */
export function renameInterfacePrefixes(
    source: string,
    interfaceName: string,
    newName: string,
): string {
    const lines = source.split(/\r?\n/);
    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    for (const hit of interfacePrefixHits(source, interfaceName)) {
        const line = lines[hit.line];
        lines[hit.line] = line.slice(0, hit.column) + newName + line.slice(hit.column + hit.length);
    }
    return lines.join(eol);
}
