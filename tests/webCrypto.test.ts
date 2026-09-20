import { describe, expect, it } from 'vitest';
import * as nodeCrypto from 'crypto';
import { createHash, randomBytes } from '../src/util/webCrypto';

// node:crypto is the oracle. These digests name module content tokens and
// backup files, so a web build that computed them differently from the
// desktop build would disagree about whether a module had changed.

const MESSAGES = [
	'',
	'a',
	'abc',
	'Module1',
	'Attribute VB_Name = "Sheet1"',
	'Модуль',
	'\u{1F600}',
	'Sub Foo()\r\n    Debug.Print 1\r\nEnd Sub\r\n',
	// Lengths around the 64-byte block and the 56-byte padding boundary,
	// which is where a hash implementation goes wrong if it goes wrong.
	'x'.repeat(54),
	'x'.repeat(55),
	'x'.repeat(56),
	'x'.repeat(57),
	'x'.repeat(63),
	'x'.repeat(64),
	'x'.repeat(65),
	'x'.repeat(119),
	'x'.repeat(120),
	'x'.repeat(128),
	'x'.repeat(1000),
];

describe.each(['sha256', 'sha1'] as const)('%s', (algorithm) => {
	it('matches node for every message', () => {
		for (const message of MESSAGES) {
			expect(
				createHash(algorithm).update(message).digest('hex'),
				`${algorithm} of ${message.length} chars`,
			).toBe(nodeCrypto.createHash(algorithm).update(message).digest('hex'));
		}
	});

	it('matches node at every byte length up to two blocks', () => {
		for (let length = 0; length <= 130; length++) {
			const bytes = Buffer.alloc(length);
			for (let i = 0; i < length; i++) {
				bytes[i] = (i * 31 + 7) & 0xff;
			}
			expect(createHash(algorithm).update(bytes).digest('hex'), `${length} bytes`).toBe(
				nodeCrypto.createHash(algorithm).update(bytes).digest('hex'),
			);
		}
	});

	it('matches node on binary data with every byte value', () => {
		const bytes = Buffer.alloc(4096);
		for (let i = 0; i < bytes.length; i++) {
			bytes[i] = i & 0xff;
		}
		expect(createHash(algorithm).update(bytes).digest('hex')).toBe(
			nodeCrypto.createHash(algorithm).update(bytes).digest('hex'),
		);
	});

	it('matches node across several update() calls', () => {
		const parts = ['Attribute ', 'VB_Name', ' = ', '"Module1"'];
		const web = createHash(algorithm);
		const node = nodeCrypto.createHash(algorithm);
		for (const part of parts) {
			web.update(part);
			node.update(part);
		}
		expect(web.digest('hex')).toBe(node.digest('hex'));
	});

	it('honours the utf8 encoding argument the callers pass', () => {
		expect(createHash(algorithm).update('Модуль', 'utf8').digest('hex')).toBe(
			nodeCrypto.createHash(algorithm).update('Модуль', 'utf8').digest('hex'),
		);
	});

	it('produces the same bytes with no encoding, and base64', () => {
		const message = 'Sheet1';
		expect([...createHash(algorithm).update(message).digest()]).toEqual([
			...nodeCrypto.createHash(algorithm).update(message).digest(),
		]);
		expect(createHash(algorithm).update(message).digest('base64')).toBe(
			nodeCrypto.createHash(algorithm).update(message).digest('base64'),
		);
	});
});

describe('the known answers', () => {
	it('reproduces the published digests of "abc"', () => {
		expect(createHash('sha256').update('abc').digest('hex')).toBe(
			'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
		);
		expect(createHash('sha1').update('abc').digest('hex')).toBe(
			'a9993e364706816aba3e25717850c26c9cd0d89d',
		);
	});
});

describe('randomBytes', () => {
	it('returns the requested length and varies', () => {
		expect(randomBytes(0).length).toBe(0);
		expect(randomBytes(32).length).toBe(32);
		expect(randomBytes(32).equals(randomBytes(32))).toBe(false);
	});
});

describe('unsupported algorithms', () => {
	it('names what it has rather than returning a wrong digest', () => {
		expect(() => createHash('md5')).toThrow(/sha256 and sha1/);
	});
});
