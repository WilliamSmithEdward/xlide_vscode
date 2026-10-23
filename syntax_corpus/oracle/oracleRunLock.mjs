// One VBE oracle run at a time per Windows user (issue #1438).
//
// Every run drives the same user's Office and VBE automation and changes the
// same user's message-box sounds, so two at once break each other. The lock
// is a named pipe rather than a file: only one process can listen on a pipe
// name, the operating system frees it the moment that process ends however
// it ends, and the owner answers who it is when a second run asks. There is
// no record to go stale, be half written or name a reused process id, and
// nothing to delete after a crash.
//
// The name comes from the user's account, through os.userInfo(), which reads
// the account rather than HOME or USERPROFILE, so a run from another
// directory, repository or environment finds the same lock. On other
// platforms the same code listens on a Unix socket path; that is for this
// module's own tests, since the oracle itself runs only on Windows.

import crypto from 'node:crypto';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';

/** How long a second run waits for the owner to say who it is. */
const OWNER_QUERY_MS = 2000;

/** The lock's address for the current account: a pipe name, or a socket path off Windows. */
export function oracleRunLockAddress(user = os.userInfo()) {
    const account = `${user.username}|${user.homedir}`;
    const id = crypto.createHash('sha256').update(account.toLowerCase()).digest('hex').slice(0, 16);
    const name = `xlide-vbe-oracle-${String(user.username).toLowerCase().replace(/[^a-z0-9._-]/g, '_')}-${id}`;
    return process.platform === 'win32' ? `\\\\.\\pipe\\${name}` : path.join(os.tmpdir(), `${name}.sock`);
}

/** What this run tells a second one about itself. */
export function oracleRunOwner(argv) {
    return {
        pid: process.pid,
        machine: os.hostname(),
        startedAt: new Date().toISOString(),
        cwd: process.cwd(),
        command: ['run_vbe_oracle.mjs', ...argv].join(' '),
    };
}

function listen(address) {
    return new Promise((resolve) => {
        const server = net.createServer();
        server.once('error', (error) => resolve({ error }));
        server.listen(address, () => resolve({ server }));
    });
}

/**
 * What the owner says it is, or why it could not be told: a refused or
 * missing socket (`absent`), or no complete record in time (`uncertain`).
 */
function queryOwner(address, timeoutMs) {
    return new Promise((resolve) => {
        let text = '';
        let settled = false;
        const finish = (result) => {
            if (settled) { return; }
            settled = true;
            clearTimeout(timer);
            socket.destroy();
            resolve(result);
        };
        const socket = net.connect(address);
        const timer = setTimeout(() => finish({ state: 'uncertain', detail: `the owner did not answer within ${timeoutMs} ms` }), timeoutMs);
        socket.setEncoding('utf8');
        socket.on('data', (chunk) => {
            text += chunk;
            const end = text.indexOf('\n');
            if (end < 0) { return; }
            try {
                const owner = JSON.parse(text.slice(0, end));
                finish(owner && typeof owner === 'object' && Number.isInteger(owner.pid)
                    ? { state: 'owned', owner }
                    : { state: 'uncertain', detail: 'the owner answered with a record that names no process' });
            } catch {
                finish({ state: 'uncertain', detail: 'the owner answered with a record that is not JSON' });
            }
        });
        socket.on('end', () => finish({ state: 'uncertain', detail: 'the owner closed the connection before a complete record' }));
        socket.on('error', (error) => finish(error.code === 'ECONNREFUSED' || error.code === 'ENOENT'
            ? { state: 'absent', detail: error.code }
            : { state: 'uncertain', detail: `asking the owner failed: ${error.code ?? error.message}` }));
    });
}

/**
 * Takes the lock at `address` for `owner` (a JSON record naming at least
 * `pid`). Resolves to `{ held: true, release }`, or `{ held: false, reason,
 * owner?, detail? }` when another live run holds it (`busy`) or ownership
 * cannot be told (`uncertain`). Another owner's lock is never removed: only
 * a Unix socket file nothing listens on is cleared, and only once.
 */
export async function acquireOracleRunLock(address, owner, { queryTimeoutMs = OWNER_QUERY_MS } = {}) {
    const record = `${JSON.stringify(owner)}\n`;
    for (let attempt = 0; attempt < 2; attempt++) {
        const { server, error } = await listen(address);
        if (server) {
            server.on('connection', (socket) => { socket.on('error', () => {}); socket.end(record); });
            // The lock never keeps the process alive; ending the process frees it.
            server.unref();
            let released = false;
            return {
                held: true,
                release: () => new Promise((resolve) => {
                    if (released) { resolve(); return; }
                    released = true;
                    server.close(() => resolve());
                }),
            };
        }
        if (error.code !== 'EADDRINUSE') {
            return { held: false, reason: 'uncertain', detail: `listening on the lock failed: ${error.code ?? error.message}` };
        }
        const answer = await queryOwner(address, queryTimeoutMs);
        if (answer.state === 'owned') {
            return { held: false, reason: 'busy', owner: answer.owner };
        }
        // A Unix socket outlives a killed owner as a file nothing listens on.
        // A Windows pipe never does, so there `absent` stays uncertain.
        if (answer.state === 'absent' && process.platform !== 'win32' && attempt === 0) {
            try { fs.unlinkSync(address); } catch { /* another contender cleared it first */ }
            continue;
        }
        return { held: false, reason: 'uncertain', detail: answer.detail };
    }
    return { held: false, reason: 'uncertain', detail: 'the lock was taken while a stale one was cleared' };
}

/** The message a refused run prints, with what to do about it. */
export function describeRefusal(result) {
    if (result.reason === 'busy') {
        const o = result.owner;
        const where = o.machine && o.machine !== os.hostname() ? ` on ${o.machine}` : '';
        return [
            `Another VBE oracle run owns this user's Office automation: process ${o.pid}${where}, started ${o.startedAt ?? 'at an unknown time'},`,
            `  in ${o.cwd ?? 'an unknown directory'}, running ${o.command ?? 'an unknown command'}.`,
            'Wait for it to finish, or stop that process. The lock is freed the moment it exits.',
        ].join('\n');
    }
    return [
        `Could not tell whether another VBE oracle run is active: ${result.detail}. Nothing was started.`,
        'If no oracle run is active, end any node process left running run_vbe_oracle.mjs and try again.',
        'The lock is a named pipe that Windows frees when its owner exits, so there is no file to delete.',
    ].join('\n');
}
