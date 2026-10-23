// The Office instances a VBE oracle run started, kept until each is known to
// be gone (issue #1439).
//
// A host is known by its image, its process id and its start time, so a
// reused process id is never taken for it and nobody else's Office session
// is touched: before anything is ended, the process at that id must still
// have that image and that start time. A record is dropped only once its
// process is gone, or is found to be someone else's.
//
// The records are also kept in a file in the user's profile. A run that was
// killed outright cannot end its own hosts; the next run, which holds the
// per-user run lock (oracleRunLock.mjs), reads the file and ends them first.
//
// Everything here is synchronous, so that an interrupt handler can use it.

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

/** How long an ended host may take to disappear before cleanup counts as failed. */
const GONE_WITHIN_MS = 10_000;

/** The per-user file of hosts a run owns, beside the sound settings backup. */
export function ownedHostsPath() {
    let home = '';
    try { home = os.userInfo().homedir; } catch { /* no account record */ }
    return path.join(home || os.homedir(), '.xlide-vbe-oracle-hosts.json');
}

function pause(ms) {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/** Windows itself: a liveness test, a process's image and start time, and taskkill. */
export const windowsHostSystem = {
    alive(pid) {
        try {
            process.kill(pid, 0);
            return true;
        } catch (error) {
            return error.code === 'EPERM';
        }
    },
    /** `{ image, start }` for the process at `pid`, `null` when there is none, `undefined` when it cannot be read. */
    identity(pid) {
        const script = `$p = Get-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue; if ($p) { $p.ProcessName.ToUpperInvariant() + ' ' + $p.StartTime.ToUniversalTime().ToFileTimeUtc() } else { 'GONE' }`;
        const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { encoding: 'utf8', timeout: 15_000 });
        const out = String(run.stdout ?? '').trim();
        if (run.status !== 0 || !out) { return undefined; }
        if (out === 'GONE') { return null; }
        const [image, start] = out.split(/\s+/);
        return image && start ? { image, start } : undefined;
    },
    end(pid, image) {
        spawnSync('taskkill', ['/PID', String(pid), '/FI', `IMAGENAME eq ${image}.EXE`, '/T', '/F'], { encoding: 'utf8', timeout: 10_000 });
    },
    pause,
    now: () => Date.now(),
};

const keyOf = (record) => `${record.pid}:${record.start}`;

export class OwnedHosts {
    /** `file` is undefined for tests that keep the records in memory only. */
    constructor({ file, system = windowsHostSystem } = {}) {
        this.file = file;
        this.system = system;
        this.records = new Map();
    }

    /** The records a previous run left in the file, now this run's to end. */
    load() {
        if (!this.file) { return []; }
        let saved;
        try {
            saved = JSON.parse(fs.readFileSync(this.file, 'utf8'));
        } catch (error) {
            if (error.code === 'ENOENT') { return []; }
            // A file that cannot be read names nothing that can be verified.
            fs.renameSync(this.file, `${this.file}.unreadable-${Date.now()}`);
            return [];
        }
        const left = Array.isArray(saved?.hosts) ? saved.hosts.filter(isRecord) : [];
        for (const record of left) { this.records.set(keyOf(record), record); }
        return left;
    }

    /** Records a host the run started. A record without a start time cannot be verified, so it is refused. */
    add(record) {
        if (!isRecord(record) || this.records.has(keyOf(record))) { return; }
        this.records.set(keyOf(record), record);
        this.save();
    }

    list() {
        return [...this.records.values()];
    }

    /**
     * Ends every recorded host but `keep`, the instance the worker still uses.
     * Returns the records that are still running and could not be ended.
     */
    sweep(keep) {
        const failed = [];
        for (const record of this.list()) {
            if (keep && keyOf(keep) === keyOf(record)) { continue; }
            if (!this.end(record)) { failed.push(record); }
        }
        return failed;
    }

    /** Ends one host if it is still the recorded one; true once it is gone or proven someone else's. */
    end(record) {
        const system = this.system;
        if (system.alive(record.pid)) {
            const now = system.identity(record.pid);
            if (now === undefined) {
                // Neither ours for certain nor gone: keep it, and say so.
                return false;
            }
            if (now !== null && now.image === record.image && now.start === record.start) {
                system.end(record.pid, record.image);
                const deadline = system.now() + GONE_WITHIN_MS;
                while (system.alive(record.pid) && system.now() < deadline) { system.pause(100); }
                if (system.alive(record.pid) && sameIdentity(system.identity(record.pid), record)) {
                    return false;
                }
            }
        }
        this.records.delete(keyOf(record));
        this.save();
        return true;
    }

    save() {
        if (!this.file) { return; }
        const hosts = this.list();
        if (hosts.length === 0) {
            fs.rmSync(this.file, { force: true });
            return;
        }
        const temp = `${this.file}.${process.pid}.tmp`;
        fs.writeFileSync(temp, `${JSON.stringify({ hosts }, null, 2)}\n`, 'utf8');
        fs.renameSync(temp, this.file);
    }
}

function isRecord(value) {
    return value !== null && typeof value === 'object' && Number.isInteger(value.pid) && value.pid > 0
        && typeof value.image === 'string' && /^[A-Z]+$/.test(value.image) && typeof value.start === 'string' && /^\d+$/.test(value.start);
}

function sameIdentity(now, record) {
    return now !== null && now !== undefined && now.image === record.image && now.start === record.start;
}

/** A line for the user about hosts that are still running, with what to do. */
export function describeLeftHosts(records) {
    const list = records.map((r) => `${r.image}.EXE pid ${r.pid}`).join(', ');
    return `Could not end ${records.length === 1 ? 'an Office instance' : 'Office instances'} this oracle run started: ${list}. `
        + 'End them in Task Manager before the next run; the next run also tries again.';
}
