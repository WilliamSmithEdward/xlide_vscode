import { runPowerShell } from './util/powershell';

export type ExcelComAvailabilityState = 'installed' | 'missing' | 'blocked' | 'unknown';

/** The Office applications the test host can drive; the probe checks the one
 * the file's container belongs to. */
export type ComProbeHostApp = 'excel' | 'word' | 'powerpoint' | 'access';

const PROBE_HOSTS: Record<ComProbeHostApp, { progId: string; noun: string }> = {
    excel: { progId: 'Excel.Application', noun: 'Excel' },
    word: { progId: 'Word.Application', noun: 'Word' },
    powerpoint: { progId: 'PowerPoint.Application', noun: 'PowerPoint' },
    access: { progId: 'Access.Application', noun: 'Access' },
};

export interface ExcelComAvailabilityStatus {
    state: ExcelComAvailabilityState;
    title: string;
    description: string;
    canRun: boolean;
}

const EXCEL_COM_PROBE_TIMEOUT_MS = 4000;

export function excelComProbePowerShellScript(hostApp: ComProbeHostApp = 'excel'): string {
    return [
        '$ErrorActionPreference = "Stop"',
        `$type = [type]::GetTypeFromProgID("${PROBE_HOSTS[hostApp].progId}")`,
        'if ($null -eq $type) { [Console]::Out.WriteLine("XLIDE_EXCEL_COM_MISSING"); exit 2 }',
        '[Console]::Out.WriteLine("XLIDE_EXCEL_COM_OK")',
    ].join('; ');
}

export function excelComAvailabilityFromProbe(
    platform: NodeJS.Platform,
    exitCode: number | null,
    stdout: string,
    stderr: string,
    hostApp: ComProbeHostApp = 'excel',
): ExcelComAvailabilityStatus {
    const noun = PROBE_HOSTS[hostApp].noun;
    if (platform !== 'win32') {
        return {
            state: 'blocked',
            title: `${noun} COM Unavailable`,
            description: `VBA tests require Microsoft ${noun} COM automation on Windows.`,
            canRun: false,
        };
    }

    if (exitCode === 0 && /XLIDE_EXCEL_COM_OK/.test(stdout)) {
        return {
            state: 'installed',
            title: `${noun} COM Ready`,
            description: `Microsoft ${noun} is registered for COM automation on this machine.`,
            canRun: true,
        };
    }

    if (exitCode === 2 || /XLIDE_EXCEL_COM_MISSING/.test(stdout)) {
        return {
            state: 'missing',
            title: `${noun} COM Not Found`,
            description: `Install Microsoft ${noun} before running VBA tests through XLIDE.`,
            canRun: false,
        };
    }

    const detail = stderr.trim() || stdout.trim() || `PowerShell exited with code ${exitCode ?? 'unknown'}.`;
    return {
        state: 'unknown',
        title: `${noun} COM Check Failed`,
        description: `XLIDE could not confirm Microsoft ${noun} COM availability: ${detail}`,
        canRun: false,
    };
}

export async function checkExcelComAvailability(
    platform: NodeJS.Platform = process.platform,
    hostApp: ComProbeHostApp = 'excel',
): Promise<ExcelComAvailabilityStatus> {
    const noun = PROBE_HOSTS[hostApp].noun;
    if (platform !== 'win32') {
        return excelComAvailabilityFromProbe(platform, null, '', '', hostApp);
    }

    const probe = await runPowerShell({
        args: ['-Command', excelComProbePowerShellScript(hostApp)],
        timeoutMs: EXCEL_COM_PROBE_TIMEOUT_MS,
    }).result;
    if (probe.timedOut) {
        return {
            state: 'unknown',
            title: `${noun} COM Check Timed Out`,
            description: `XLIDE could not confirm Microsoft ${noun} COM availability before the setup check timed out.`,
            canRun: false,
        };
    }
    if (probe.spawnError) {
        return {
            state: 'unknown',
            title: `${noun} COM Check Failed`,
            description: `XLIDE could not run the ${noun} COM availability check: ${probe.spawnError.message}`,
            canRun: false,
        };
    }
    return excelComAvailabilityFromProbe(
        platform,
        probe.code,
        probe.stdoutLines.join('\n'),
        probe.stderrLines.join('\n'),
        hostApp,
    );
}
