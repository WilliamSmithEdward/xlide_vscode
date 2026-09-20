import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { ZipArchive } from '../src/vba/zip';
import { containerCodec, setContainerCodec } from '../src/vba/containerCodec';
import { platformCodec as nodeCodec } from '../src/vba/containerCodecNode';
import { platformCodec as webCodec } from '../src/vba/containerCodecWeb';
import { openMacroContainer } from '../src/vba/macroContainer';

// The web extension host runs the whole container engine through
// containerCodecWeb. These hold the two codecs to the same bytes on real
// containers, which is the difference between a browser build that works and
// one that quietly writes files Office will not open.

const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'binaries');

/** OOXML packages: the ZIP-based containers, which is where the codec runs. */
const ZIP_FIXTURES = fs
	.readdirSync(FIXTURE_DIR)
	.filter((f) => /\.(xlsm|xltm|xlam|xlsb|docm|dotm|pptm|ppsm|ppam)$/i.test(f))
	.sort();

function withCodec<T>(codec: typeof nodeCodec, run: () => T): T {
	setContainerCodec(codec);
	try {
		return run();
	} finally {
		setContainerCodec(nodeCodec);
	}
}

afterEach(() => setContainerCodec(nodeCodec));

describe('container codecs agree', () => {
	it('has fixtures to run against', () => {
		expect(ZIP_FIXTURES.length).toBeGreaterThan(5);
	});

	it('defaults to zlib so desktop builds and tests are unchanged', () => {
		expect(containerCodec().name).toBe('zlib');
	});

	it.each(ZIP_FIXTURES)('reads every entry of %s identically', (file) => {
		const bytes = fs.readFileSync(path.join(FIXTURE_DIR, file));

		const viaNode = withCodec(nodeCodec, () => readAll(bytes));
		const viaWeb = withCodec(webCodec, () => readAll(bytes));

		expect([...viaWeb.keys()].sort()).toEqual([...viaNode.keys()].sort());
		for (const [name, expected] of viaNode) {
			const actual = viaWeb.get(name);
			expect(actual, `${file} -> ${name} missing`).toBeDefined();
			expect(actual!.equals(expected), `${file} -> ${name} differs`).toBe(true);
		}
		expect(viaNode.size).toBeGreaterThan(0);
	});

	it('opens every VBA project identically', () => {
		// Some fixtures are cells-and-formulas only; they exercise the ZIP
		// path above but have no project to open.
		let compared = 0;
		for (const file of ZIP_FIXTURES) {
			const bytes = fs.readFileSync(path.join(FIXTURE_DIR, file));
			const viaNode = withCodec(nodeCodec, () => projectDirBytes(bytes));
			if (!viaNode) {
				continue;
			}
			const viaWeb = withCodec(webCodec, () => projectDirBytes(bytes));
			expect(viaWeb, `${file} opened under zlib but not under the web codec`).toBeDefined();
			expect(viaWeb!.equals(viaNode), `${file} project differs`).toBe(true);
			compared++;
		}
		expect(compared).toBeGreaterThanOrEqual(8);
	});
});

describe('entries written by the browser codec', () => {
	const file = 'ShapesFixture.xlsm';
	const target = 'xl/workbook.xml';

	it('are stored, and read back through zlib as what was written', () => {
		const original = fs.readFileSync(path.join(FIXTURE_DIR, file));
		const payload = Buffer.from('<written-by-the-web-build/>'.repeat(200), 'utf8');

		const saved = withCodec(webCodec, () => {
			const archive = ZipArchive.read(original);
			archive.write(target, payload);
			return archive.toBytes();
		});

		// Reopened through the normal desktop path, with no knowledge that a
		// browser wrote it.
		const reopened = ZipArchive.read(saved);
		expect(reopened.read(target).equals(payload)).toBe(true);
	});

	it('leave every untouched entry byte-identical', () => {
		const original = fs.readFileSync(path.join(FIXTURE_DIR, file));
		const before = readAll(original);

		const saved = withCodec(webCodec, () => {
			const archive = ZipArchive.read(original);
			archive.write(target, Buffer.from('<changed/>', 'utf8'));
			return archive.toBytes();
		});

		const after = readAll(saved);
		expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
		for (const [name, expected] of before) {
			if (name === target) {
				continue;
			}
			expect(after.get(name)!.equals(expected), `${name} was perturbed`).toBe(true);
		}
	});

	it('cost only the rewritten entry its compression', () => {
		const original = fs.readFileSync(path.join(FIXTURE_DIR, file));
		const payload = Buffer.from('<a>'.repeat(4000), 'utf8');

		const viaWeb = withCodec(webCodec, () => {
			const archive = ZipArchive.read(original);
			archive.write(target, payload);
			return archive.toBytes();
		});
		const viaNode = withCodec(nodeCodec, () => {
			const archive = ZipArchive.read(original);
			archive.write(target, payload);
			return archive.toBytes();
		});

		// The stored entry is bigger, but bounded by that one part: every
		// other entry carried over compressed.
		expect(viaWeb.length).toBeGreaterThan(viaNode.length);
		expect(viaWeb.length - viaNode.length).toBeLessThanOrEqual(payload.length);
	});
});

describe('what the browser codec cannot do', () => {
	it('refuses to write a PowerPoint compressed storage, with a reason', () => {
		expect(() => webCodec.deflate(Buffer.from('x'))).toThrow(/not supported in the browser/);
	});
});

/** Every entry of a package, expanded. */
function readAll(bytes: Buffer): Map<string, Buffer> {
	const archive = ZipArchive.read(bytes);
	const out = new Map<string, Buffer>();
	for (const name of archive.names()) {
		if (name.endsWith('/')) {
			continue; // directory entry, no payload
		}
		out.set(name, archive.read(name));
	}
	return out;
}

/**
 * The VBA project's `dir` stream: proof the CFB underneath expanded too, not
 * just the ZIP around it. Undefined when the package carries no project.
 */
function projectDirBytes(bytes: Buffer): Buffer | undefined {
	let cfb;
	try {
		cfb = openMacroContainer(bytes).vbaCfb();
	} catch {
		return undefined;
	}
	return cfb.hasStreamInStorage('VBA', 'dir')
		? cfb.getStreamInStorage('VBA', 'dir')
		: cfb.getStream('dir');
}
