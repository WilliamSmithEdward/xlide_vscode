// `readFolderAnnotation` walks the text a line at a time instead of splitting
// it, because splitting a 27,000-line module to look at its first four lines
// cost 0.467 ms per module on every listing and now costs 0.001 ms. That is a
// hand-optimized loop over the plainest possible one, so this pins the two
// together over generated input rather than over the handful of cases anyone
// thinks to write down.

import { describe, expect, it } from 'vitest';
import { normalizeFolderPath, readFolderAnnotation } from '../src/vba/folderAnnotation';

const PROCEDURE_HEADER_RE =
	/^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+/i;
const FOLDER_TAG_RE = /@folder(?=[\s()"]|$)/i;

/**
 * The obvious implementation: split, then look at each line. Deliberately not
 * clever, and deliberately not sharing the walk under test.
 */
function reference(text: string, truncated = false): { folder?: string; complete: boolean } {
	const lines = text.split(/\r\n|\r|\n/);
	const limit = truncated && lines[lines.length - 1] !== '' ? lines.length - 1 : lines.length;
	for (let i = 0; i < limit; i += 1) {
		const line = lines[i];
		if (PROCEDURE_HEADER_RE.test(line)) {
			return { complete: true };
		}
		const comment = commentOf(line);
		if (comment === undefined) { continue; }
		const tag = FOLDER_TAG_RE.exec(comment);
		if (!tag) { continue; }
		const folder = normalizeFolderPath(argumentOf(comment.slice(tag.index + tag[0].length)));
		return folder ? { folder, complete: true } : { complete: true };
	}
	return { complete: false };
}

function commentOf(line: string): string | undefined {
	let inString = false;
	for (let i = 0; i < line.length; i += 1) {
		const ch = line[i];
		if (ch === '"') { inString = !inString; continue; }
		if (inString) { continue; }
		if (ch === "'") { return line.slice(i + 1); }
		if (/^rem(?=[\s:]|$)/i.test(line.slice(i)) && /^[ \t]*$/.test(line.slice(0, i))) {
			return line.slice(i + 3);
		}
	}
	return undefined;
}

function argumentOf(rest: string): string {
	const bracketed = /^[ \t]*\([ \t]*"([^"]*)"?/.exec(rest);
	if (bracketed) { return bracketed[1]; }
	const quoted = /^[ \t]*"([^"]*)"?/.exec(rest);
	if (quoted) { return quoted[1]; }
	const bare = /^[ \t]*\(?([^)]*)\)?/.exec(rest);
	return bare ? bare[1] : '';
}

/** A tiny deterministic generator, so a failure is reproducible from its seed. */
function makeRandom(seed: number): () => number {
	let state = seed >>> 0;
	return () => {
		state = (state * 1664525 + 1013904223) >>> 0;
		return state / 0x100000000;
	};
}

const LINES = [
	'',
	'   ',
	'Option Explicit',
	'Public Total As Long',
	"'@Folder(\"Accounts.Ledger\")",
	"'@Folder Accounts",
	"'@folder(accounts.ledger)",
	"'@Folder \"A.B\"",
	"'@Folders(\"NotThis\")",
	"'@Folder-ish",
	"'@Folder",
	"'@Folder(\"\")",
	"' just a comment",
	'Rem @Folder("FromRem")',
	'Rem plain',
	'Public Const S As String = "it\'s @Folder(Wrong)"',
	'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
	'Sub T()',
	'    Debug.Print 1',
	'End Sub',
	'Public Function F() As Long',
	'Property Get Name() As String',
	'Friend Sub S()',
	'Private Static Function G()',
	'Attribute VB_Name = "Mod1"',
	'VERSION 1.0 CLASS',
	'   \'@Folder("Indented")',
	'x = "unclosed string',
];

const EOLS = ['\r\n', '\n', '\r'];

describe('the walked reader answers what the split one would', () => {
	it('agrees on 4000 generated modules, whole and truncated', () => {
		const random = makeRandom(20260905);
		const disagreements: string[] = [];
		for (let n = 0; n < 4000; n += 1) {
			const count = 1 + Math.floor(random() * 8);
			const parts: string[] = [];
			for (let i = 0; i < count; i += 1) {
				parts.push(LINES[Math.floor(random() * LINES.length)]);
			}
			const eol = EOLS[Math.floor(random() * EOLS.length)];
			// Half the samples end mid-line, which is what a prefix read gives.
			const text = parts.join(eol) + (random() < 0.5 ? eol : '');
			for (const truncated of [false, true]) {
				const got = readFolderAnnotation(text, { truncated });
				const want = reference(text, truncated);
				if (got.folder !== want.folder || got.complete !== want.complete) {
					disagreements.push(
						`${JSON.stringify(text)} truncated=${truncated}: `
						+ `${JSON.stringify(got)} vs ${JSON.stringify(want)}`,
					);
				}
			}
		}
		expect(disagreements.slice(0, 5)).toEqual([]);
	});
});
