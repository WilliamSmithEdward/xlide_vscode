import { describe, expect, it } from 'vitest';
import {
    normalizeVbaTestSupportModuleSource,
    XLIDE_ASSERT_MODULE_NAME,
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
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNothing');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub IsNotNothing');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub Fail');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub Throws');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub DoesNotThrow');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Application.Run macroName');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Err.Raise XLIDE_ASSERTION_ERROR');
    });

    it('normalizes line endings before comparing installed module source', () => {
        expect(normalizeVbaTestSupportModuleSource('Option Explicit\r\n\r\n')).toBe('Option Explicit');
        expect(normalizeVbaTestSupportModuleSource('Option Explicit\n')).toBe('Option Explicit');
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
