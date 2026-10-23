// The VBE oracle's per-user run lock (issue #1438), in isolation: each test
// takes its own address, so nothing here touches a real run's lock.

import { spawn } from 'node:child_process';
import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as net from 'node:net';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { acquireOracleRunLock, describeRefusal, oracleRunLockAddress } from '../syntax_corpus/oracle/oracleRunLock.mjs';

type Lock = { held: true; release: () => Promise<void> } | { held: false; reason: 'busy' | 'uncertain'; owner?: Record<string, unknown>; detail?: string };
const acquire = acquireOracleRunLock as (address: string, owner: object, options?: { queryTimeoutMs?: number }) => Promise<Lock>;

const MODULE = path.join(__dirname, '..', 'syntax_corpus', 'oracle', 'oracleRunLock.mjs');
const cleanups: Array<() => Promise<void> | void> = [];
afterEach(async () => {
	while (cleanups.length) {
		await cleanups.pop()!();
	}
});

function freshAddress(): string {
	const name = `xlide-oracle-lock-test-${crypto.randomBytes(6).toString('hex')}`;
	return process.platform === 'win32' ? `\\\\.\\pipe\\${name}` : path.join(os.tmpdir(), `${name}.sock`);
}

async function held(address: string, owner: object = { pid: process.pid }): Promise<() => Promise<void>> {
	const lock = await acquire(address, owner);
	if (!lock.held) {
		throw new Error(`expected the lock, got ${lock.reason}`);
	}
	cleanups.push(lock.release);
	return lock.release;
}

/** A server at `address` that answers each connection with `reply`, or never, for an impostor or a stuck owner. */
async function responder(address: string, reply: string | undefined): Promise<void> {
	const server = net.createServer((socket) => {
		if (reply !== undefined) {
			socket.end(reply);
		}
	});
	await new Promise<void>((resolve) => server.listen(address, resolve));
	cleanups.push(() => new Promise<void>((resolve) => server.close(() => resolve())));
}

describe('VBE oracle run lock', () => {
	it('refuses a second run while the first holds it, and names the owner', async () => {
		const address = freshAddress();
		const release = await held(address, { pid: 4242, startedAt: 'then', cwd: 'C:\\work', command: 'run_vbe_oracle.mjs --case x' });
		const second = await acquire(address, { pid: 1 });
		expect(second.held).toBe(false);
		expect(second).toMatchObject({ reason: 'busy', owner: { pid: 4242, cwd: 'C:\\work' } });
		const message = describeRefusal(second) as string;
		expect(message).toContain('process 4242');
		expect(message).toContain('run_vbe_oracle.mjs --case x');
		await release();
		const third = await acquire(address, { pid: 2 });
		expect(third.held).toBe(true);
		if (third.held) { cleanups.push(third.release); }
	});

	it('lets exactly one of several simultaneous runs take it', async () => {
		const address = freshAddress();
		const results = await Promise.all(Array.from({ length: 8 }, (_, i) => acquire(address, { pid: 1000 + i })));
		const winners = results.filter((r) => r.held);
		expect(winners).toHaveLength(1);
		for (const winner of winners) {
			if (winner.held) { cleanups.push(winner.release); }
		}
		expect(results.filter((r) => !r.held).every((r) => !r.held && r.reason === 'busy')).toBe(true);
	});

	it('is free again once its owner is killed, with nothing left to delete', async () => {
		const address = freshAddress();
		const child = spawn(process.execPath, ['--input-type=module', '-e', [
			`const { acquireOracleRunLock } = await import(${JSON.stringify(new URL(`file:///${MODULE.replace(/\\/g, '/')}`).href)});`,
			`const lock = await acquireOracleRunLock(${JSON.stringify(address)}, { pid: process.pid });`,
			"process.stdout.write(lock.held ? 'HELD\\n' : 'NOT\\n');",
			'setInterval(() => {}, 1000);',
		].join('\n')], { stdio: ['ignore', 'pipe', 'inherit'] });
		cleanups.push(() => { child.kill('SIGKILL'); });
		const first = await new Promise<string>((resolve) => child.stdout.once('data', (chunk) => resolve(String(chunk).trim())));
		expect(first).toBe('HELD');
		expect(await acquire(address, { pid: 1 })).toMatchObject({ held: false, reason: 'busy', owner: { pid: child.pid } });
		child.kill('SIGKILL');
		await new Promise((resolve) => child.once('exit', resolve));
		// A Unix socket file is left behind; nothing listens on it, so it is stale.
		const after = await acquire(address, { pid: 2 });
		expect(after.held).toBe(true);
		if (after.held) { cleanups.push(after.release); }
	});

	it('refuses, and keeps, an owner whose record is partial, not JSON or names no process', async () => {
		for (const reply of ['{"pid": 12', 'not json at all\n', '{"cwd": "C:\\\\x"}\n']) {
			const address = freshAddress();
			await responder(address, reply);
			const result = await acquire(address, { pid: 1 });
			expect(result, reply).toMatchObject({ held: false, reason: 'uncertain' });
			expect(describeRefusal(result) as string).toContain('Nothing was started');
			// The other owner still listens: it was not removed.
			expect(await acquire(address, { pid: 1 })).toMatchObject({ held: false });
		}
	});

	it('refuses an owner that does not answer, after the query timeout', async () => {
		const address = freshAddress();
		await responder(address, undefined);
		const started = Date.now();
		const result = await acquire(address, { pid: 1 }, { queryTimeoutMs: 200 });
		expect(result).toMatchObject({ held: false, reason: 'uncertain' });
		expect(Date.now() - started).toBeLessThan(2000);
	});

	it('names the owner\'s machine when it is another', () => {
		const message = describeRefusal({ held: false, reason: 'busy', owner: { pid: 7, machine: 'ELSEWHERE-PC' } }) as string;
		expect(message).toContain('process 7 on ELSEWHERE-PC');
		expect(describeRefusal({ held: false, reason: 'busy', owner: { pid: 7, machine: os.hostname() } }) as string).not.toContain(' on ');
	});

	it('takes its address from the account, whatever HOME and USERPROFILE say', () => {
		const before = oracleRunLockAddress() as string;
		const saved = { HOME: process.env.HOME, USERPROFILE: process.env.USERPROFILE };
		try {
			process.env.HOME = 'C:\\elsewhere';
			process.env.USERPROFILE = 'C:\\elsewhere';
			expect(oracleRunLockAddress()).toBe(before);
		} finally {
			for (const [key, value] of Object.entries(saved)) {
				if (value === undefined) { delete process.env[key]; } else { process.env[key] = value; }
			}
		}
		expect(oracleRunLockAddress({ username: 'Ann', homedir: 'C:\\Users\\Ann' })).not.toBe(oracleRunLockAddress({ username: 'Bob', homedir: 'C:\\Users\\Bob' }));
		expect(oracleRunLockAddress({ username: 'Ann', homedir: 'C:\\Users\\Ann' })).toBe(oracleRunLockAddress({ username: 'ANN', homedir: 'c:\\users\\ann' }));
	});

	it('costs little to take and to refuse', async () => {
		const address = freshAddress();
		let t0 = performance.now();
		await held(address);
		const take = performance.now() - t0;
		t0 = performance.now();
		await acquire(address, { pid: 1 });
		const refuse = performance.now() - t0;
		// Measured at a few milliseconds each; the bound only catches a hang.
		expect(take).toBeLessThan(500);
		expect(refuse).toBeLessThan(500);
	});

	it('is taken before the coordinator starts a worker or silences dialog sounds', () => {
		const coordinator = fs.readFileSync(path.join(__dirname, '..', 'syntax_corpus', 'oracle', 'run_vbe_oracle.mjs'), 'utf8');
		const main = coordinator.slice(coordinator.indexOf('async function main('));
		const lock = main.indexOf('await acquireOracleRunLock(oracleRunLockAddress()');
		expect(lock).toBeGreaterThan(0);
		expect(lock).toBeLessThan(main.indexOf('new OracleWorker('));
		expect(lock).toBeLessThan(main.indexOf("setDialogSounds('silence')"));
		expect(main.slice(lock, main.indexOf('new OracleWorker('))).toContain('return 3;');
	});
});
