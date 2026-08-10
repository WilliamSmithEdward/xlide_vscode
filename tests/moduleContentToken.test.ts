import { describe, expect, it } from 'vitest';
import { checkModuleContentToken, moduleContentToken } from '../src/moduleContentToken';

// xlide_writeModule takes the FULL module source, so the blast radius of a
// stale read is the whole module rather than the lines being changed. Nothing
// carried a version, so a change Excel or a user made after an agent's read
// was overwritten silently.
const SOURCE = 'Option Explicit\r\n\r\nPublic Sub Go()\r\n    Debug.Print 1\r\nEnd Sub\r\n';

describe('module content tokens', () => {
	it('is stable for the same content', () => {
		expect(moduleContentToken(SOURCE)).toBe(moduleContentToken(SOURCE));
	});

	it('changes when the content changes', () => {
		expect(moduleContentToken(SOURCE))
			.not.toBe(moduleContentToken(SOURCE.replace('Debug.Print 1', 'Debug.Print 2')));
	});

	it('ignores line-ending style', () => {
		// The engine writes CRLF and callers often hold LF; that round trip is
		// not somebody else's edit.
		expect(moduleContentToken(SOURCE)).toBe(moduleContentToken(SOURCE.replace(/\r\n/g, '\n')));
	});

	it('is recognisable in a transcript', () => {
		expect(moduleContentToken(SOURCE)).toMatch(/^xlide1:[0-9a-f]{32}$/);
	});
});

describe('a conditional write', () => {
	it('proceeds when the module is unchanged', () => {
		expect(checkModuleContentToken(SOURCE, moduleContentToken(SOURCE), 'M')).toBeUndefined();
	});

	it('is refused when the module moved under the caller', () => {
		const readAt = moduleContentToken(SOURCE);
		const changedSince = SOURCE.replace('Debug.Print 1', 'Debug.Print 999');
		const rejection = checkModuleContentToken(changedSince, readAt, 'M');
		expect(rejection).toBeDefined();
		expect(rejection?.message).toContain('changed since it was read');
		// The message has to carry the current token, or the caller cannot retry.
		expect(rejection?.message).toContain(moduleContentToken(changedSince));
	});

	it('stays opt-in when no token is supplied', () => {
		// Existing callers keep working; the guard is something you ask for.
		expect(checkModuleContentToken(SOURCE, undefined, 'M')).toBeUndefined();
	});
});
