import { runPowerShell } from './util/powershell';
import { OFFICE_HOST_APPS as PROBE_HOSTS, type OfficeHostApp } from './officeHostApps';

export type OfficeComAvailabilityState = 'installed' | 'missing' | 'blocked' | 'unknown';

/** The Office applications the test host can drive; the probe checks the one
 * the file's container belongs to. */
export type ComProbeHostApp = OfficeHostApp;

export interface OfficeComAvailabilityStatus {
    state: OfficeComAvailabilityState;
    title: string;
    description: string;
    canRun: boolean;
}

const OFFICE_COM_PROBE_TIMEOUT_MS = 4000;

export function officeComProbePowerShellScript(hostApp: ComProbeHostApp = 'excel'): string {
    return [
        '$ErrorActionPreference = "Stop"',
        `$type = [type]::GetTypeFromProgID("${PROBE_HOSTS[hostApp].progId}")`,
        'if ($null -eq $type) { [Console]::Out.WriteLine("XLIDE_OFFICE_COM_MISSING"); exit 2 }',
        '[Console]::Out.WriteLine("XLIDE_OFFICE_COM_OK")',
    ].join('\n');
}

export function officeComAvailabilityFromProbe(
    platform: NodeJS.Platform,
    exitCode: number | null,
    stdout: string,
    stderr: string,
    hostApp: ComProbeHostApp = 'excel',
): OfficeComAvailabilityStatus {
    const noun = PROBE_HOSTS[hostApp].noun;
    if (platform !== 'win32') {
        return {
            state: 'blocked',
            title: `${noun} COM Unavailable`,
            description: `VBA tests require Microsoft ${noun} COM automation on Windows.`,
            canRun: false,
        };
    }

    if (exitCode === 0 && /XLIDE_OFFICE_COM_OK/.test(stdout)) {
        return {
            state: 'installed',
            title: `${noun} COM Ready`,
            description: `Microsoft ${noun} is registered for COM automation on this machine.`,
            canRun: true,
        };
    }

    if (exitCode === 2 || /XLIDE_OFFICE_COM_MISSING/.test(stdout)) {
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

export async function checkOfficeComAvailability(
    platform: NodeJS.Platform = process.platform,
    hostApp: ComProbeHostApp = 'excel',
): Promise<OfficeComAvailabilityStatus> {
    const noun = PROBE_HOSTS[hostApp].noun;
    if (platform !== 'win32') {
        return officeComAvailabilityFromProbe(platform, null, '', '', hostApp);
    }

    const probe = await runPowerShell({
        script: officeComProbePowerShellScript(hostApp),
        timeoutMs: OFFICE_COM_PROBE_TIMEOUT_MS,
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
    return officeComAvailabilityFromProbe(
        platform,
        probe.code,
        probe.stdoutLines.join('\n'),
        probe.stderrLines.join('\n'),
        hostApp,
    );
}
