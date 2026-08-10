// Content tokens for agent reads and writes.
//
// An agent reads a module, reasons about it for a while, then writes the whole
// module back - `xlide_writeModule` takes the full source, so the blast radius
// of a stale read is the entire module, not the lines being changed. Nothing
// carried a version, so anything Excel, a user, or another agent changed in
// between was overwritten silently.
//
// A token is a hash of the module source as read. Passing it back on write
// makes the write conditional: same token, the module is untouched and the
// write proceeds; different token, it is rejected and the caller re-reads.

import { createHash } from 'crypto';

/** Prefix so a token is recognisable in a tool transcript. */
const TOKEN_PREFIX = 'xlide1:';

/** Content token for a module's source, stable across line-ending style. */
export function moduleContentToken(source: string): string {
    // Normalized so a CRLF/LF difference on the round trip is not mistaken for
    // someone else's edit - the engine writes CRLF, callers often hold LF.
    const normalized = source.replace(/\r\n/g, '\n');
    return TOKEN_PREFIX + createHash('sha256').update(normalized, 'utf8').digest('hex').slice(0, 32);
}

export interface StaleWriteRejection {
    expected: string;
    actual: string;
    message: string;
}

/**
 * Returns a rejection when `expectedToken` does not describe the module's
 * current content. Undefined means the write may proceed - including when the
 * caller supplied no token, which keeps the guard opt-in.
 */
export function checkModuleContentToken(
    currentSource: string,
    expectedToken: string | undefined,
    moduleName: string,
): StaleWriteRejection | undefined {
    if (!expectedToken) {
        return undefined;
    }
    const actual = moduleContentToken(currentSource);
    if (actual === expectedToken) {
        return undefined;
    }
    return {
        expected: expectedToken,
        actual,
        message:
            `Module "${moduleName}" changed since it was read, so the write was refused `
            + 'to avoid discarding that change. Read the module again, reapply your edit to '
            + `the current source, and write with the new contentToken (${actual}).`,
    };
}
