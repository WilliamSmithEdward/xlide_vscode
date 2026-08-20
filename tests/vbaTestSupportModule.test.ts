import { describe, expect, it } from 'vitest';
import {
    normalizeVbaTestSupportModuleSource,
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_REVISION,
    XLIDE_ASSERT_MODULE_SOURCE,
} from '../src/vbaTestSupportModule';

describe('VBA test support module', () => {
    it('exports a standard-module assertion API callable as XlideAssert.*', () => {
        expect(XLIDE_ASSERT_MODULE_NAME).toBe('XlideAssert');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Attribute VB_Name = "XlideAssert"');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub AreEqual');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub AreNotEqual');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsTrue');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsFalse');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub AreSame');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub AreNotSame');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNothing');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNotNothing');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNullValue');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNotNullValue');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsEmptyValue');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNotEmptyValue');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub Contains');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub DoesNotContain');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub StartsWith');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub EndsWith');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub Fail');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub WriteLine');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub Throws');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub DoesNotThrow');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Application.Run MacroName');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Private mLastFailureMessage As String');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Private mOutputJsonItems As String');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub ResetLastFailure');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub ResetOutput');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub ResetTestState');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Function LastFailureMessage() As String');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Function OutputJson() As String');
        expect(XLIDE_ASSERT_MODULE_SOURCE).not.toContain('Public Function RunTest');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Private Function JsonEscape');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Chr$(92) & Chr$(34)');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('If Len(mLastFailureMessage) = 0 Then');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('mLastFailureMessage = Message');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Private Function TryString');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('vbBinaryCompare');
        expect(XLIDE_ASSERT_MODULE_SOURCE).not.toContain('Err.Raise XLIDE_ASSERTION_ERROR');
    });

    it('normalizes line endings and identifier case before comparing installed source', () => {
        // Identifiers fold (VBA owns their casing); the rest is verbatim.
        expect(normalizeVbaTestSupportModuleSource('Option Explicit\r\n\r\n')).toBe('option explicit');
        expect(normalizeVbaTestSupportModuleSource('Option Explicit\n')).toBe('option explicit');
        expect(normalizeVbaTestSupportModuleSource('Debug.Print "Kept As Written" \' And So Is This'))
            .toBe('debug.print "Kept As Written" \' And So Is This');
    });

    it('stamps a revision so a casing-only fix still reaches installed copies (issue #38)', () => {
        // Case folding would read revision 1 - identical but for the lowercase
        // parameters - as already installed. The stamp is a comment, and
        // comments compare verbatim, so the revision is what carries.
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain(
            `' XLIDE test support module, revision ${XLIDE_ASSERT_MODULE_REVISION}.`,
        );
        const previousRevision = XLIDE_ASSERT_MODULE_SOURCE
            .replace(`revision ${XLIDE_ASSERT_MODULE_REVISION}`, `revision ${XLIDE_ASSERT_MODULE_REVISION - 1}`)
            .replace(/\bMessage\b/g, 'message');
        expect(normalizeVbaTestSupportModuleSource(previousRevision)).not.toBe(
            normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE),
        );
    });

    it('reads a module someone else re-cased as installed, not outdated (issue #38)', () => {
        // Another declaration in the project - a user's own `Dim value` - can
        // re-case XlideAssert's identifiers under it. VBA treats that as the
        // same code, and so must the version gate.
        const reCased = XLIDE_ASSERT_MODULE_SOURCE
            .replace(/\bMessage\b/g, 'message')
            .replace(/\bValue As Variant\b/g, 'value As Variant');
        expect(reCased).not.toBe(XLIDE_ASSERT_MODULE_SOURCE);
        expect(normalizeVbaTestSupportModuleSource(reCased)).toBe(
            normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE),
        );
    });

    it('still reads edited text and string casing as a different module', () => {
        // Case folding is for identifiers, which VBA re-cases; a changed string
        // literal or a changed line is a real difference and must still show.
        const editedLiteral = XLIDE_ASSERT_MODULE_SOURCE.replace('"Expected True but was False."', '"expected true but was false."');
        expect(editedLiteral).not.toBe(XLIDE_ASSERT_MODULE_SOURCE);
        expect(normalizeVbaTestSupportModuleSource(editedLiteral)).not.toBe(
            normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE),
        );

        const editedCode = `${XLIDE_ASSERT_MODULE_SOURCE}\r\nPublic Sub Extra()\r\nEnd Sub\r\n`;
        expect(normalizeVbaTestSupportModuleSource(editedCode)).not.toBe(
            normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE),
        );
    });

    it('treats the workbook-visible body as matching the bundled support module', () => {
        const visibleBody = XLIDE_ASSERT_MODULE_SOURCE
            .replace(/\r\n|\r/g, '\n')
            .split('\n')
            .filter((line) => !/^Attribute\s+VB_/i.test(line))
            .join('\n');

        expect(normalizeVbaTestSupportModuleSource(visibleBody)).toBe(
            normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE),
        );
    });
});
