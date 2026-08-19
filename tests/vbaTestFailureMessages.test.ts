import { describe, expect, it } from 'vitest';
import { vbaTestFailureMessage } from '../src/vbaTestFailureMessages';

describe('VBA test failure messages', () => {
    it('strips host protocol prefixes and keeps plain messages intact', () => {
        expect(vbaTestFailureMessage(new Error('RUN_FAILED|Assertion failed'))).toBe('Assertion failed');
        expect(vbaTestFailureMessage('OPEN_FAILED|XLIDE could not open the workbook read-only for tests: locked'))
            .toBe('XLIDE could not open the workbook read-only for tests: locked');
        expect(vbaTestFailureMessage('RUNNER_FAILED|spawn powershell.exe ENOENT'))
            .toBe('spawn powershell.exe ENOENT');
        expect(vbaTestFailureMessage('A user message with | punctuation')).toBe('A user message with | punctuation');
    });

    it('maps the raw macro-run HRESULT emitted by the host script to friendly guidance', () => {
        expect(vbaTestFailureMessage(new Error([
            'RUN_FAILED|Exception calling "Run" with "1" argument(s):',
            '"Exception from HRESULT: 0x800A9C68"',
            'HRESULT: 0x80131501',
            'Inner: Exception from HRESULT: 0x800A9C68',
            'Inner HRESULT: 0x800A9C68',
            'At C:\\Users\\William\\AppData\\Local\\Temp\\xlide-vba-test-host-DDT6Nq\\run-vba-tests.ps1:677 char:4587',
            '+ ... tPrefix, $excelId, $macroName) }; $excel.Run($macroRef);',
            '+                       ~~~~~~~~~~~~~~~~~~~~',
        ].join('\n')))).toBe(
            'The Office application could not run the test macro. Check for VBA compile errors, macro security prompts, or a missing test procedure. HRESULT: 0x800A9C68.',
        );
    });

    it('maps the raw RPC-failure HRESULTs emitted by the host script to friendly guidance', () => {
        expect(vbaTestFailureMessage([
            'RUN_FAILED|HRESULT: 0x80131501',
            'Exception calling "Run" with "2" argument(s): "The remote procedure call failed. (Exception from HRESULT: 0x800706BE)"',
            'Inner: The remote procedure call failed. (Exception from HRESULT: 0x800706BE)',
            'Inner HRESULT: 0x800706BE',
        ].join('\n'))).toBe(
            'The Office application hosting the tests became unavailable while running them. It may have closed, crashed, or been blocked by a modal dialog. HRESULT: 0x800706BE.',
        );
        expect(vbaTestFailureMessage([
            'RUN_FAILED|HRESULT: 0x80131501',
            'Exception calling "Run" with "2" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)"',
            'Inner HRESULT: 0x800706BA',
        ].join('\n'))).toBe(
            'The Office application hosting the tests became unavailable while running them. It may have closed, crashed, or been blocked by a modal dialog. HRESULT: 0x800706BA.',
        );
    });
});
