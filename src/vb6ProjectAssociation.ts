// Whether Windows can open a VB6 project file at all.
//
// F5 on a VB6 project hands the `.vbp` to the shell, which is Visual Basic
// itself where it is installed. On a machine where nothing claims the
// extension the shell fails with "Application not found" in a dialog of its
// own, which says nothing about what XLIDE was doing or what the user can do
// instead. So F5 asks first, and says it itself.
//
// The answer comes from the registry, read through `reg.exe`: the user's own
// choice for the extension, then the machine's default, then the "Open with"
// candidates Windows offers when there is no default. Nothing here knows or
// names any particular application; it reports what the machine says.

import { spawnSync } from 'child_process';

export interface Vb6ProjectAssociation {
	/** The shell has a handler: opening the project will work. */
	opensDirectly: boolean;
	/**
	 * What Windows offers under "Open with" when no default is set, as
	 * executable names taken from each candidate's own open command.
	 */
	candidates: string[];
	/** Not Windows, or the registry could not be read: F5 should just try. */
	unknown: boolean;
}

/** The ProgId a `reg query ... /ve` prints, or undefined when the value is unset. */
export function progIdFromDefaultValue(stdout: string): string | undefined {
	const match = /^\s*\(Default\)\s+REG_\w+\s+(.+?)\s*$/im.exec(stdout);
	const value = match?.[1];
	return value && value !== '(value not set)' ? value : undefined;
}

/** The ProgId a `reg query ... /v ProgId` prints. */
export function progIdFromNamedValue(stdout: string, name: string): string | undefined {
	const match = new RegExp(`^\\s*${name}\\s+REG_\\w+\\s+(.+?)\\s*$`, 'im').exec(stdout);
	return match?.[1] || undefined;
}

/** Every ProgId listed under a key's values, which is how OpenWithProgids records them. */
export function progIdsFromValueNames(stdout: string): string[] {
	const out: string[] = [];
	for (const line of stdout.split(/\r?\n/)) {
		const match = /^\s{4,}(\S+)\s+REG_\w+/.exec(line);
		if (match && match[1] !== '(Default)') {
			out.push(match[1]);
		}
	}
	return out;
}

/** The executable a shell open command runs, by its file name alone. */
export function executableFromOpenCommand(command: string | undefined): string | undefined {
	if (!command) {
		return undefined;
	}
	const quoted = /^"([^"]+)"/.exec(command.trim());
	const raw = quoted ? quoted[1] : command.trim().split(/\s+/)[0];
	const name = raw?.split(/[\\/]/).pop();
	return name || undefined;
}

function query(...args: string[]): string | undefined {
	try {
		const result = spawnSync('reg', ['query', ...args], { encoding: 'utf8', windowsHide: true });
		return result.status === 0 ? result.stdout : undefined;
	} catch {
		return undefined;
	}
}

function openCommandFor(progId: string): string | undefined {
	const stdout = query(`HKEY_CLASSES_ROOT\\${progId}\\shell\\open\\command`, '/ve');
	return stdout ? progIdFromDefaultValue(stdout) : undefined;
}

let cached: Vb6ProjectAssociation | undefined;

/** What the machine can do with a `.vbp`, read once per session. */
export function vb6ProjectAssociation(platform: string = process.platform): Vb6ProjectAssociation {
	if (platform !== 'win32') {
		return { opensDirectly: false, candidates: [], unknown: true };
	}
	if (cached) {
		return cached;
	}
	const userChoice = query(
		'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.vbp\\UserChoice',
		'/v', 'ProgId',
	);
	const machine = query('HKEY_CLASSES_ROOT\\.vbp', '/ve');
	if (userChoice === undefined && machine === undefined) {
		// The registry did not answer at all; do not claim it cannot open.
		cached = { opensDirectly: false, candidates: [], unknown: true };
		return cached;
	}
	const chosen = (userChoice ? progIdFromNamedValue(userChoice, 'ProgId') : undefined)
		?? (machine ? progIdFromDefaultValue(machine) : undefined);
	if (chosen && openCommandFor(chosen)) {
		cached = { opensDirectly: true, candidates: [], unknown: false };
		return cached;
	}
	const offered = query('HKEY_CLASSES_ROOT\\.vbp\\OpenWithProgids');
	const candidates = (offered ? progIdsFromValueNames(offered) : [])
		.map((progId) => executableFromOpenCommand(openCommandFor(progId)))
		.filter((name): name is string => name !== undefined);
	cached = { opensDirectly: false, candidates: [...new Set(candidates)], unknown: false };
	return cached;
}

/** Forgets the reading, for a test. */
export function resetVb6ProjectAssociationForTests(): void {
	cached = undefined;
}
