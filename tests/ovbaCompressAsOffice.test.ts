import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { compress, decompress } from '../src/vba/ovba';
import { openMacroContainer } from '../src/vba/macroContainer';
import { VbaProject } from '../src/vba/vbaProject';
import type { Cfb } from '../src/vba/cfb';

// Issue #149, from pyOpenVBA 6.3.0: Office chooses its copy tokens as the
// LZNT1 engine in ntdll does, and writes one extra empty flag byte when a
// chunk's last token fills its flag byte. compress() follows the same rule,
// so a stream Office wrote compresses back to exactly its own bytes.

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

const hex = (b: Buffer): string => [...b].map((x) => x.toString(16).padStart(2, '0')).join(' ');

/** Every compressed dir and module stream Office wrote in the fixtures, with the file it came from. */
function officeStreams(): Array<{ label: string; stored: Buffer }> {
	const out: Array<{ label: string; stored: Buffer }> = [];
	for (const name of fs.readdirSync(FIXTURES).sort()) {
		let cfb: Cfb;
		let project: VbaProject;
		try {
			cfb = openMacroContainer(fs.readFileSync(path.join(FIXTURES, name))).vbaCfb();
			project = VbaProject.parse(cfb);
		} catch {
			continue; // no VBA project, or a container kept another way
		}
		let inStorage = true;
		try { cfb.getStreamInStorage('VBA', 'dir'); } catch { inStorage = false; }
		const read = (stream: string): Buffer => (inStorage ? cfb.getStreamInStorage('VBA', stream) : cfb.getStream(stream));
		out.push({ label: `${name}/dir`, stored: read('dir') });
		for (const module of project.modules) {
			out.push({ label: `${name}/${module.name}`, stored: read(module.streamName).subarray(module.textOffset) });
		}
	}
	return out;
}

describe('compress() writes the bytes Office writes (issue #149)', () => {
	it('encodes the issue\'s two small cases', () => {
		expect(hex(compress(Buffer.from('ABCDEFGH')))).toBe('01 09 b0 00 41 42 43 44 45 46 47 48 00');
		expect(hex(compress(Buffer.from('ABCDEFGABC')))).toBe('01 0a b0 80 41 42 43 44 45 46 47 00 60 00');
	});

	it('compresses every Office-written stream in the fixtures back to its own bytes', () => {
		const streams = officeStreams();
		// Excel, Word and PowerPoint, .xls, .xlsb, .doc and .ppt among them.
		expect(streams.length).toBeGreaterThan(100);
		const differing = streams
			.filter(({ label, stored }) => !compress(decompress(stored, label)).equals(stored))
			.map(({ label }) => label);
		expect(differing).toEqual([]);
	});

	it('still round-trips input with no repeats, a full chunk, and more than one chunk', () => {
		const random = Buffer.alloc(9000);
		let seed = 12345;
		for (let i = 0; i < random.length; i++) {
			seed = (seed * 1103515245 + 12345) >>> 0;
			random[i] = seed >>> 24;
		}
		const text = Buffer.from('Attribute VB_Name = "Module1"\r\n'.repeat(400));
		for (const input of [Buffer.alloc(0), Buffer.from('A'), Buffer.from('AB'), random, random.subarray(0, 4096), text]) {
			expect(decompress(compress(input))).toEqual(input);
		}
	});
});
