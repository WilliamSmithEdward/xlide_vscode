// The VBA grammar in a real VS Code: the TextMate scopes VS Code computes for
// a loose module, read through the command its own colorizer tests use.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, workspaceRoot } from './support';

interface SyntaxToken {
	/** The token's text. */
	c: string;
	/** Its scopes, space-separated. */
	t: string;
}

/** The scopes of the first token holding `text` in `file`, as VS Code colors it. */
async function scopesOf(file: string, text: string): Promise<string> {
	const tokens = await vscode.commands.executeCommand<SyntaxToken[]>(
		'_workbench.captureSyntaxTokens', vscode.Uri.file(file),
	);
	const token = (tokens ?? []).find((candidate) => candidate.c.includes(text));
	assert.ok(token, `no token holds ${JSON.stringify(text)}`);
	return token.t;
}

suite('VBA grammar', () => {
	suiteSetup(async () => {
		await activate();
	});

	test('colors the line a comment ending in _ runs on to as comment text (issue #82)', async () => {
		const file = path.join(workspaceRoot(), 'GrammarProbe.bas');
		fs.writeFileSync(file, [
			'Sub Test()',
			"    ' disabled: _",
			'    End Sub',
			'    Debug.Print 1',
			"#If True Then ' note _",
			'    not code',
			'#End If',
			'    x = a + _   ',
			'        b',
			'End Sub',
			'',
		].join('\r\n'));
		try {
			assert.match(await scopesOf(file, 'End Sub'), /\bcomment\.line\.apostrophe\.vba\b/);
			assert.match(await scopesOf(file, 'not code'), /\bcomment\.line\.apostrophe\.vba\b/);
			assert.doesNotMatch(await scopesOf(file, 'Debug'), /comment/);
			assert.match(await scopesOf(file, ' _   '), /\bpunctuation\.separator\.continuation\.vba\b/);
		} finally {
			fs.rmSync(file, { force: true });
		}
	});
});
