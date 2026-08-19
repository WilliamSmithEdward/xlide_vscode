import { normalizeEol } from './vbaSourceScan';
import { errorMessage } from './util/errors';

// Single owner of the user-facing failure text for VBA test runs: strips the
// host protocol prefixes, maps known Excel automation HRESULTs to friendly
// guidance, and drops PowerShell stack noise. The embedded host script emits
// raw HRESULT and exception text only.
export function vbaTestFailureMessage(error: unknown): string {
    const raw = errorMessage(error);
    const unwrapped = raw.replace(/^(?:RUN_FAILED|OPEN_FAILED|RUNNER_FAILED|TIMEOUT|HOST_ERROR)\|/, '');
    return cleanVbaTestFailureMessage(unwrapped);
}

function cleanVbaTestFailureMessage(raw: string): string {
    const text = normalizeEol(raw).trim();
    if (!text) {
        return text;
    }
    const rpcHresult = /0x800706(?:BE|BA)/i.exec(text)?.[0];
    if (rpcHresult && /Exception calling "Run"|RPC server|remote procedure call|HRESULT:/i.test(text)) {
        return `The Office application hosting the tests became unavailable while running them. It may have closed, crashed, or been blocked by a modal dialog. HRESULT: 0x${rpcHresult.slice(2).toUpperCase()}.`;
    }
    if (/0x800A9C68/i.test(text) && /Exception calling "Run"|Exception from HRESULT|run-vba-tests\.ps1|HRESULT:/i.test(text)) {
        return 'The Office application could not run the test macro. Check for VBA compile errors, macro security prompts, or a missing test procedure. HRESULT: 0x800A9C68.';
    }
    const lines = text
        .split('\n')
        .map((line) => line.trim())
        .filter((line) => line.length > 0 && !isNoisyPowerShellDetail(line));
    if (lines.length === 0) {
        return text;
    }
    return [...new Set(lines)].join('\n');
}

function isNoisyPowerShellDetail(line: string): boolean {
    return /^At .*run-vba-tests\.ps1:\d+ char:\d+/i.test(line) ||
        /^\+ /.test(line) ||
        /^~{6,}$/.test(line) ||
        /^CategoryInfo\s*:/i.test(line) ||
        /^FullyQualifiedErrorId\s*:/i.test(line);
}
