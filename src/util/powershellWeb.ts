// The PowerShell spawner's web twin: there is no shell in a browser, and no
// local Office for one to talk to.
//
// Nothing in the web build should reach this. The callers - the Office
// launcher, the COM availability probe, the write coordinator, the test host
// - are all behind platformFeatures or are only entered when a host app is
// present, which in a browser it never is. Throwing rather than resolving
// with a failure keeps a wrong assumption loud instead of turning into a
// silently skipped coordination step.

import type { PowerShellRun, RunPowerShellOptions } from './powershell';

export function runPowerShell(_options: RunPowerShellOptions): PowerShellRun {
    throw new Error(
        'XLIDE in the browser cannot run PowerShell. This feature needs the desktop editor.',
    );
}
