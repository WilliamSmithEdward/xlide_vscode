import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { setHostPlatform } from '../src/vba/hostPlatform';
import { platformHost as nodeHost } from '../src/vba/hostPlatformNode';
import {
	HostFileNotPrimedError,
	forgetHostFiles,
	platformHost as webHost,
	primeHostFile,
	primeHostFileAbsent,
	takeHostFileWrites,
} from '../src/vba/hostPlatformWeb';
import {
	listModules,
	readModule,
	resetProjectCacheForTests,
	writeModule,
} from '../src/vba/projectService';

// The web extension host runs the engine against hostPlatformWeb's in-memory
// filesystem, primed by the extension layer before each call and drained
// after. These assert it behaves like the real one where the engine can tell
// the difference - and, at the end, that a workbook edited entirely through
// it is a workbook the ordinary desktop path reads back.

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'ShapesFixture.xlsm');

/** Loads a file into the web platform the way the extension layer will. */
function prime(filePath: string): void {
	const stat = fs.statSync(filePath);
	primeHostFile(filePath, fs.readFileSync(filePath), stat.mtimeMs);
}

function withWebPlatform<T>(run: () => T): T {
	forgetHostFiles();
	takeHostFileWrites();
	setHostPlatform(webHost);
	resetProjectCacheForTests();
	try {
		return run();
	} finally {
		setHostPlatform(nodeHost);
		resetProjectCacheForTests();
	}
}

afterEach(() => {
	setHostPlatform(nodeHost);
	forgetHostFiles();
	takeHostFileWrites();
	resetProjectCacheForTests();
});

describe('the engine on the in-memory platform', () => {
	it('lists the same modules as it does on disk', () => {
		resetProjectCacheForTests();
		const onDisk = listModules(FIXTURE).map((m) => m.name).sort();

		const inMemory = withWebPlatform(() => {
			prime(FIXTURE);
			return listModules(FIXTURE).map((m) => m.name).sort();
		});

		expect(inMemory).toEqual(onDisk);
		expect(onDisk.length).toBeGreaterThan(0);
	});

	it('reads every module identically', () => {
		resetProjectCacheForTests();
		const names = listModules(FIXTURE).map((m) => m.name);
		const onDisk = new Map(names.map((n) => [n, readModule(FIXTURE, n).source]));

		const inMemory = withWebPlatform(() => {
			prime(FIXTURE);
			return new Map(names.map((n) => [n, readModule(FIXTURE, n).source]));
		});

		for (const name of names) {
			expect(inMemory.get(name), `${name} differs`).toBe(onDisk.get(name));
		}
		// Document modules are legitimately empty; at least one module in the
		// fixture must carry code or this compares nothing.
		expect([...onDisk.values()].some((s) => s.length > 0)).toBe(true);
	});

	it('mints distinct random bytes', () => {
		const a = webHost.randomBytes(16);
		const b = webHost.randomBytes(16);
		expect(a.length).toBe(16);
		expect(a.equals(b)).toBe(false);
	});
});

describe('priming', () => {
	it('refuses a path nobody loaded, rather than reading it as empty', () => {
		withWebPlatform(() => {
			expect(() => webHost.readFile(FIXTURE)).toThrow(HostFileNotPrimedError);
			expect(() => webHost.stat(FIXTURE)).toThrow(HostFileNotPrimedError);
			expect(() => webHost.exists(FIXTURE)).toThrow(HostFileNotPrimedError);
		});
	});

	it('reports a file primed as absent the way the real platform does', () => {
		const missing = path.join(__dirname, 'fixtures', 'binaries', 'NoSuchFile.frx');
		expect(nodeHost.exists(missing)).toBe(false);
		expect(nodeHost.statIfPresent(missing)).toBeUndefined();
		expect(nodeHost.readFileIfPresent(missing)).toBeUndefined();

		withWebPlatform(() => {
			primeHostFileAbsent(missing);
			expect(webHost.exists(missing)).toBe(false);
			expect(webHost.statIfPresent(missing)).toBeUndefined();
			expect(webHost.readFileIfPresent(missing)).toBeUndefined();
		});
	});

	it('reports the stat the workspace gave it, which the parse cache keys on', () => {
		const stat = fs.statSync(FIXTURE);
		withWebPlatform(() => {
			prime(FIXTURE);
			expect(webHost.stat(FIXTURE)).toEqual({ mtimeMs: stat.mtimeMs, size: stat.size });
		});
	});
});

describe('a workbook edited entirely in memory', () => {
	it('comes back as a workbook the ordinary desktop path reads', () => {
		resetProjectCacheForTests();
		const name = listModules(FIXTURE)[0].name;
		const source = `Attribute VB_Name = "${name}"\r\nSub WrittenInTheBrowser()\r\n    Debug.Print "hello from the web"\r\nEnd Sub\r\n`;

		const written = withWebPlatform(() => {
			prime(FIXTURE);
			const result = writeModule(FIXTURE, name, source);
			expect(result.ok).toBe(true);

			// What the extension layer would flush through workspace.fs.
			const writes = takeHostFileWrites();
			expect(writes.size).toBe(1);
			return writes.get(FIXTURE)!;
		});

		// Land it as a real file and read it back with no web pieces involved.
		const out = path.join(
			__dirname,
			'fixtures',
			'binaries',
			`.hostPlatform-roundtrip-${process.pid}.xlsm`,
		);
		try {
			fs.writeFileSync(out, written);
			resetProjectCacheForTests();
			expect(readModule(out, name).source).toContain('WrittenInTheBrowser');
			expect(listModules(out).map((m) => m.name).sort()).toEqual(
				listModules(FIXTURE).map((m) => m.name).sort(),
			);
		} finally {
			fs.rmSync(out, { force: true });
		}
	});

	it('leaves the file on disk untouched', () => {
		const before = fs.readFileSync(FIXTURE);
		withWebPlatform(() => {
			prime(FIXTURE);
			writeModule(FIXTURE, listModules(FIXTURE)[0].name, 'Sub Nothing()\r\nEnd Sub\r\n');
			takeHostFileWrites();
		});
		expect(fs.readFileSync(FIXTURE).equals(before)).toBe(true);
	});
});
