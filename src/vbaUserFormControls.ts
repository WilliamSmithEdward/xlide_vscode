// Controls declared by a UserForm's designer.
//
// A form's controls are members of the form's class, declared by the designer
// rather than by any line of code, so code-behind that says `RegionPick.AddItem`
// is correct VBA. The analyzer resolves a form like an ordinary class, where
// those declarations do not exist, and reported every reference as an
// undeclared variable - five findings on a small real form.
//
// The declarations are not out of reach: a `.frm` header carries the control
// tree in the same `Begin ... End` blocks the module-sync header scan already
// walks. This reads them out, so the names (and their types) can be handed to
// the analyzer as implicit members of the form.
//
// Kept free of any `vscode` import so it can be unit-tested directly.

/** A control the designer declared on a form. */
export interface UserFormControl {
    /** Name the code-behind uses, e.g. `RegionPick`. */
    name: string;
    /** Programmatic id exactly as the designer wrote it, e.g. `Forms.ComboBox.1`. */
    progId: string;
    /** Type a member lookup should resolve against, e.g. `MSForms.ComboBox`. */
    type: string;
}

/**
 * `Begin <progId> <name>` opens a control block; the outermost one is the form
 * itself and is not a control of itself.
 */
// The class id is taken as an opaque token: a form's own is a `{...}` GUID and
// a control's is a prog id, and being strict about what is inside the braces
// would only reject headers this has no need to understand.
const BEGIN_RE = /^\s*Begin\s+(\{[^}]*\}|[\w.]+)\s+([\p{L}_][\p{L}\p{M}\p{N}_]*)/u;
const END_RE = /^\s*End\s*$/i;

/**
 * Reads the controls a `.frm` designer header declares. Returns an empty list
 * for source that is not a form header, so callers can pass any module.
 */
export function parseUserFormControls(source: string): UserFormControl[] {
    const lines = source.split(/\r?\n/);
    let index = 0;
    while (index < lines.length && lines[index].trim() === '') {
        index += 1;
    }
    if (index >= lines.length || !/^\s*VERSION\b/i.test(lines[index])) {
        return [];
    }

    const out: UserFormControl[] = [];
    let depth = 0;
    for (index += 1; index < lines.length; index += 1) {
        const line = lines[index];
        const begin = BEGIN_RE.exec(line);
        if (begin) {
            depth += 1;
            // depth 1 is the form; only what it contains is a control.
            if (depth > 1) {
                const progId = begin[1];
                out.push({ name: begin[2], progId, type: controlTypeFor(progId) });
            }
            continue;
        }
        if (END_RE.test(line)) {
            depth -= 1;
            if (depth <= 0) {
                break;
            }
            continue;
        }
        // Property lines and blanks are the only other thing a header holds; a
        // line that is neither means the block never closed, so stop rather
        // than run on into the code and invent controls out of it.
        if (line.trim() !== '' && !/^\s*[\w.]+\s*=/.test(line)) {
            break;
        }
    }
    return out;
}

/**
 * `Forms.ComboBox.1` is how the designer writes it; `MSForms.ComboBox` is what
 * the type library calls it. An id in a shape we do not recognise is passed
 * through, since a wrong guess is worse than an unresolved type.
 */
function controlTypeFor(progId: string): string {
    const match = /^Forms\.([A-Za-z][\w]*)(?:\.\d+)?$/.exec(progId);
    return match ? `MSForms.${match[1]}` : progId;
}
