// Diagnostics tests: a Word Range and the Selection have a default member,
// Text, and a Paragraph's is its Range (issue #438). Measured in Word 16.0
// (2026-10-01). The defaults come from the type library, DISPID 0, through
// src/analyzer/host/hostDefaultMembers.ts.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { HOST_DEFAULT_MEMBERS } from '../../src/analyzer/host/hostDefaultMembers';

const RULES = ['set-required', 'object-default-value'];

function found(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src, { host: 'word' }).filter((diag) => RULES.includes(diag.code)).map((diag) => `${diag.code}: ${diag.message}`);
}

describe('Word objects with a default member', () => {
	it('take a Let and give a value through it', () => {
		for (const [decl, set] of [['Range', 'Set x = ActiveDocument.Range(0, 0)'], ['Selection', 'Set x = Selection']]) {
			for (const use of ['x = 5', 'Main = x', 'Main = x & "a"']) {
				expect(found(`Dim x As ${decl}`, set, use), `${decl}: ${use}`).toEqual([]);
			}
		}
		expect(found('Dim x As Paragraph', 'Set x = ActiveDocument.Paragraphs(1)', 'Main = x')).toEqual([]);
	});

	it('reads the defaults the type library gives', () => {
		expect(HOST_DEFAULT_MEMBERS['Word.Range']).toMatchObject({ name: 'Text', kind: 'property', required: 0 });
		expect(HOST_DEFAULT_MEMBERS['Word.Paragraphs']).toMatchObject({ name: 'Item', required: 1 });
		expect(HOST_DEFAULT_MEMBERS['Word.Font']).toBeUndefined();
	});
});
