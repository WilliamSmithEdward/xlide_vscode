// The VBE oracle's record of the Office instances a run started (issue
// #1439), against a mocked process table: nothing here ends a real process.

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { describeLeftHosts, OwnedHosts, windowsHostSystem } from '../../syntax_corpus/oracle/oracleOwnedHosts.mjs';

interface Proc { image: string; start: string; dies?: boolean }

/** A process table: `end` removes a process unless it is marked to survive. */
function system(table: Map<number, Proc>, readable = true) {
	let clock = 0;
	const ended: number[] = [];
	return {
		ended,
		alive: (pid: number) => table.has(pid),
		identity: (pid: number) => (!readable ? undefined : table.has(pid) ? { image: table.get(pid)!.image, start: table.get(pid)!.start } : null),
		end: (pid: number) => {
			ended.push(pid);
			if (table.get(pid)?.dies !== false) { table.delete(pid); }
		},
		pause: (ms: number) => { clock += ms; },
		now: () => clock,
	};
}

const excel = { pid: 100, image: 'EXCEL', start: '133000000000000000' };

const dirs: string[] = [];
afterEach(() => {
	for (const dir of dirs.splice(0)) { fs.rmSync(dir, { recursive: true, force: true }); }
});
function tempFile(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-owned-hosts-'));
	dirs.push(dir);
	return path.join(dir, 'hosts.json');
}

describe('VBE oracle owned hosts', () => {
	it('drops a host that is already gone, ending nothing', () => {
		const sys = system(new Map());
		const owned = new OwnedHosts({ system: sys });
		owned.add(excel);
		expect(owned.sweep()).toEqual([]);
		expect(owned.list()).toEqual([]);
		expect(sys.ended).toEqual([]);
	});

	it('ends a host that is still the recorded process, and drops it once gone', () => {
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start }]]));
		const owned = new OwnedHosts({ system: sys });
		owned.add(excel);
		expect(owned.sweep()).toEqual([]);
		expect(sys.ended).toEqual([100]);
		expect(owned.list()).toEqual([]);
	});

	it('never ends a process that reused the id, by start time or by image', () => {
		for (const other of [{ image: 'EXCEL', start: '999' }, { image: 'NOTEPAD', start: excel.start }]) {
			const sys = system(new Map([[100, other]]));
			const owned = new OwnedHosts({ system: sys });
			owned.add(excel);
			expect(owned.sweep()).toEqual([]);
			expect(sys.ended).toEqual([]);
			expect(owned.list()).toEqual([]);
		}
	});

	it('keeps, and reports, a host whose identity cannot be read', () => {
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start }]]), false);
		const owned = new OwnedHosts({ system: sys });
		owned.add(excel);
		expect(owned.sweep()).toEqual([excel]);
		expect(sys.ended).toEqual([]);
		expect(owned.list()).toEqual([excel]);
	});

	it('keeps, and reports, a host that survives being ended, after the deadline', () => {
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start, dies: false }]]));
		const owned = new OwnedHosts({ system: sys });
		owned.add(excel);
		expect(owned.sweep()).toEqual([excel]);
		expect(sys.ended).toEqual([100]);
		expect(sys.now()).toBeGreaterThanOrEqual(10_000);
		expect(owned.list()).toEqual([excel]);
		expect(describeLeftHosts([excel])).toContain('EXCEL.EXE pid 100');
	});

	it('leaves the instance the worker still drives', () => {
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start }], [200, { image: 'WINWORD', start: '5' }]]));
		const owned = new OwnedHosts({ system: sys });
		const word = { pid: 200, image: 'WINWORD', start: '5' };
		owned.add(excel);
		owned.add(word);
		expect(owned.sweep(word)).toEqual([]);
		expect(sys.ended).toEqual([100]);
		expect(owned.list()).toEqual([word]);
	});

	it('refuses a record it could not verify later', () => {
		const owned = new OwnedHosts({ system: system(new Map()) });
		for (const bad of [{ pid: 1, image: 'EXCEL' }, { pid: 0, image: 'EXCEL', start: '1' }, { pid: 1, image: 'excel', start: '1' }, { pid: 1, image: 'EXCEL', start: 'x' }]) {
			owned.add(bad as never);
		}
		expect(owned.list()).toEqual([]);
	});

	it('cleans up twice harmlessly', () => {
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start }]]));
		const owned = new OwnedHosts({ system: sys });
		owned.add(excel);
		owned.sweep();
		owned.sweep();
		expect(sys.ended).toEqual([100]);
	});
});

describe('VBE oracle owned hosts on disk', () => {
	it('keeps the records in the file while any is outstanding, and removes it after', () => {
		const file = tempFile();
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start }]]));
		const owned = new OwnedHosts({ file, system: sys });
		owned.add(excel);
		expect(JSON.parse(fs.readFileSync(file, 'utf8'))).toEqual({ hosts: [excel] });
		owned.sweep();
		expect(fs.existsSync(file)).toBe(false);
	});

	it('hands a killed run\'s hosts to the next run, which ends them', () => {
		const file = tempFile();
		fs.writeFileSync(file, JSON.stringify({ hosts: [excel, { pid: 5, image: 'EXCEL' }] }));
		const sys = system(new Map([[100, { image: 'EXCEL', start: excel.start }]]));
		const next = new OwnedHosts({ file, system: sys });
		expect(next.load()).toEqual([excel]);
		expect(next.sweep()).toEqual([]);
		expect(sys.ended).toEqual([100]);
		expect(fs.existsSync(file)).toBe(false);
	});

	it('moves an unreadable file aside rather than act on it', () => {
		const file = tempFile();
		fs.writeFileSync(file, '{"hosts": [');
		const owned = new OwnedHosts({ file, system: system(new Map()) });
		expect(owned.load()).toEqual([]);
		expect(fs.existsSync(file)).toBe(false);
		expect(fs.readdirSync(path.dirname(file)).some((name) => name.startsWith('hosts.json.unreadable-'))).toBe(true);
	});
});

describe('the coordinator and worker use one ownership contract', () => {
	const oracle = (name: string): string => fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', name), 'utf8');

	it('records each host with its start time, and never forgets one that survives a stop', () => {
		const worker = oracle('vbe_oracle_worker.ps1');
		const record = worker.slice(worker.indexOf('function Write-HostRecord'), worker.indexOf('function Stop-OracleHost'));
		expect(record).toContain('.StartTime.ToUniversalTime().ToFileTimeUtc()');
		expect(record).toContain('$($script:appPid) $($script:hostImageOfApp) $start');
		const stop = worker.slice(worker.indexOf('function Stop-OracleHost'), worker.indexOf('\nfunction ', worker.indexOf('function Stop-OracleHost') + 1));
		expect(stop).toMatch(/if \(\$null -ne \$left\) \{\s*throw "The /);
		expect(stop.indexOf('throw "The ')).toBeLessThan(stop.indexOf('$script:appPid = 0'));
	});

	it('ends hosts only through the owned record, on every path', () => {
		const coordinator = oracle('run_vbe_oracle.mjs');
		expect(coordinator).not.toMatch(/function killProcess|killRecordedHost|taskkill', \[\s*'\/PID', String\(recorded/);
		expect(coordinator.match(/owned\.sweep\(/g)?.length).toBeGreaterThanOrEqual(8);
		// Before a case, before a new worker, after a timeout, after a worker exit,
		// at close, at an interrupt, at the end, and for an earlier run's hosts.
		for (const site of ['const left = worker.owned.sweep(worker.child ? worker.hostProcess : undefined);',
			'const left = this.owned.sweep();\n        if (left.length) { throw new HostCleanupError(left); }',
			'worker.kill();\n            worker.owned.sweep();',
			"process.once('exit', cleanUp);",
			'const inherited = owned.load();']) {
			expect(coordinator.replace(/\r\n/g, '\n'), site).toContain(site);
		}
	});
});

describe.runIf(process.platform === 'win32')('Windows process identity', () => {
	it('reads this process\'s image and start time, and nothing for an id with no process', () => {
		expect(windowsHostSystem.alive(process.pid)).toBe(true);
		const me = windowsHostSystem.identity(process.pid);
		expect(me?.image).toBe('NODE');
		expect(me?.start).toMatch(/^\d{17,19}$/);
		expect(windowsHostSystem.identity(process.pid)).toEqual(me);
		// Windows process ids are multiples of four; this one is never handed out.
		expect(windowsHostSystem.alive(4_294_967_292)).toBe(false);
	});
});
