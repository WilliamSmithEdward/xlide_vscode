import { runPowerShell } from './util/powershell';

export type ExcelComAvailabilityState = 'installed' | 'missing' | 'blocked' | 'unknown';

export interface ExcelComAvailabilityStatus {
    state: ExcelComAvailabilityState;
    title: string;
    description: string;
    canRun: boolean;
}

const EXCEL_COM_PROBE_TIMEOUT_MS = 4000;

export function excelComProbePowerShellScript(): string {
    return [
        '$ErrorActionPreference = "Stop"',
        '$type = [type]::GetTypeFromProgID("Excel.Application")',
        'if ($null -eq $type) { [Console]::Out.WriteLine("XLIDE_EXCEL_COM_MISSING"); exit 2 }',
        '[Console]::Out.WriteLine("XLIDE_EXCEL_COM_OK")',
    ].join('; ');
}

export function excelComAvailabilityFromProbe(
    platform: NodeJS.Platform,
    exitCode: number | null,
    stdout: string,
    stderr: string,
): ExcelComAvailabilityStatus {
    if (platform !== 'win32') {
        return {
            state: 'blocked',
            title: 'Excel COM Unavailable',
            description: 'Workbook tests require Microsoft Excel COM automation on Windows.',
            canRun: false,
        };
    }

    if (exitCode === 0 && /XLIDE_EXCEL_COM_OK/.test(stdout)) {
        return {
            state: 'installed',
            title: 'Excel COM Ready',
            description: 'Microsoft Excel is registered for COM automation on this machine.',
            canRun: true,
        };
    }

    if (exitCode === 2 || /XLIDE_EXCEL_COM_MISSING/.test(stdout)) {
        return {
            state: 'missing',
            title: 'Excel COM Not Found',
            description: 'Install Microsoft Excel before running workbook tests through XLIDE.',
            canRun: false,
        };
    }

    const detail = stderr.trim() || stdout.trim() || `PowerShell exited with code ${exitCode ?? 'unknown'}.`;
    return {
        state: 'unknown',
        title: 'Excel COM Check Failed',
        description: `XLIDE could not confirm Microsoft Excel COM availability: ${detail}`,
        canRun: false,
    };
}

export async function checkExcelComAvailability(
    platform: NodeJS.Platform = process.platform,
): Promise<ExcelComAvailabilityStatus> {
    if (platform !== 'win32') {
        return excelComAvailabilityFromProbe(platform, null, '', '');
    }

    const probe = await runPowerShell({
        args: ['-Command', excelComProbePowerShellScript()],
        timeoutMs: EXCEL_COM_PROBE_TIMEOUT_MS,
    }).result;
    if (probe.timedOut) {
        return {
            state: 'unknown',
            title: 'Excel COM Check Timed Out',
            description: 'XLIDE could not confirm Microsoft Excel COM availability before the setup check timed out.',
            canRun: false,
        };
    }
    if (probe.spawnError) {
        return {
            state: 'unknown',
            title: 'Excel COM Check Failed',
            description: `XLIDE could not run the Excel COM availability check: ${probe.spawnError.message}`,
            canRun: false,
        };
    }
    return excelComAvailabilityFromProbe(
        platform,
        probe.code,
        probe.stdoutLines.join('\n'),
        probe.stderrLines.join('\n'),
    );
}
