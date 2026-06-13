// Structural guard: every diagnostic rule function that emits diagnostics (takes
// a `push: PushFn`) must be wired into DIAGNOSTIC_RULE_REGISTRY, otherwise it is
// defined-but-never-run in production (analyzeModule only executes the registry).
//
// This catches the "added a rule + metadata + audit but forgot the registry
// entry" mistake, which tsc and corpusProvenance do not cover on their own.

import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { describe, it, expect } from 'vitest';

describe('diagnostic rule registry coverage', () => {
	it('wires every push-taking rule function into DIAGNOSTIC_RULE_REGISTRY', () => {
		const diagnosticsDir = join(process.cwd(), 'src', 'analyzer', 'diagnostics');
		const registry = readFileSync(join(diagnosticsDir, 'registry.ts'), 'utf-8');
		const rulesDir = join(diagnosticsDir, 'rules');

		const unregistered: string[] = [];
		let checked = 0;

		for (const file of readdirSync(rulesDir).filter((f) => f.endsWith('.ts'))) {
			const text = readFileSync(join(rulesDir, file), 'utf-8');
			for (const match of text.matchAll(/export function (\w+)\s*\(/g)) {
				const name = match[1];
				// Signature window: from the declaration to the body's opening brace.
				const braceIndex = text.indexOf('{', match.index ?? 0);
				const signature = braceIndex >= 0 ? text.slice(match.index ?? 0, braceIndex) : '';
				if (!/\bpush\s*:\s*PushFn\b/.test(signature)) {
					continue; // not a diagnostic-emitting rule entry point
				}
				checked += 1;
				if (!new RegExp(`\\b${name}\\b`).test(registry)) {
					unregistered.push(`${name} (rules/${file})`);
				}
			}
		}

		// Sanity: the scan actually found the rule functions (guards against a
		// regex change silently making this test vacuous).
		expect(checked).toBeGreaterThan(60);
		expect(unregistered).toEqual([]);
	});
});
