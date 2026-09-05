// Move to Module (issue #69). Qualified call sites are repointed; unqualified
// ones are left alone, because VBA resolves a bare call to a public procedure
// across the whole project - rewriting them would be a large diff that changes
// nothing. The refusal names what would be stranded rather than just declining.

import { describe, expect, it } from 'vitest';
import { moveToModule } from '../src/analyzer/refactor/moveToModule';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

const SOURCE = [
	'Option Explicit',
	'',
	'Public Sub Build()',
	'    Debug.Print "built"',
	'End Sub',
	'',
	'Public Sub Other()',
	'End Sub',
	'',
].join('\r\n');

function run(source: string, procedure: string, others: Record<string, string>, target = 'Helpers') {
	const offset = source.indexOf(procedure);
	expect(offset, `'${procedure}' is not in the fixture`).toBeGreaterThan(-1);
	return moveToModule({
		source,
		offset,
		moduleName: 'Reports',
		targetModuleName: target,
		otherModuleSources: others,
	});
}

function reason(source: string, procedure: string, others: Record<string, string>, target = 'Helpers'): string {
	const result = run(source, procedure, others, target);
	if (result.ok) { throw new Error(`expected a refusal, got ${result.title}`); }
	return result.reason;
}

describe('what it writes', () => {
	it('takes the procedure out of the module it leaves', () => {
		const result = run(SOURCE, 'Public Sub Build', { Helpers: 'Option Explicit\r\n' });
		if (!result.ok) { throw new Error(result.reason); }
		const left = applyVbaTextEdits(SOURCE, result.edits);
		expect(left).not.toContain('Build');
		expect(left).toContain('Public Sub Other()');
		expect(left).toBe('Option Explicit\r\n\r\nPublic Sub Other()\r\nEnd Sub\r\n');
	});

	it('appends it to the module it joins', () => {
		const helpers = 'Option Explicit\r\n';
		const result = run(SOURCE, 'Public Sub Build', { Helpers: helpers });
		if (!result.ok) { throw new Error(result.reason); }
		const target = result.otherModules!.find((m) => m.moduleName === 'Helpers')!;
		expect(applyVbaTextEdits(helpers, target.edits)).toBe([
			'Option Explicit',
			'',
			'Public Sub Build()',
			'    Debug.Print "built"',
			'End Sub',
			'',
		].join('\r\n'));
	});

	it('repoints a qualified call in another module', () => {
		const caller = 'Public Sub Go()\r\n    Reports.Build\r\nEnd Sub\r\n';
		const result = run(SOURCE, 'Public Sub Build', { Helpers: '', Caller: caller });
		if (!result.ok) { throw new Error(result.reason); }
		const edits = result.otherModules!.find((m) => m.moduleName === 'Caller')!;
		expect(applyVbaTextEdits(caller, edits.edits)).toContain('    Helpers.Build\r\n');
	});

	it('repoints a qualified call in the module it leaves', () => {
		const source = SOURCE + 'Public Sub Local()\r\n    Reports.Build\r\nEnd Sub\r\n';
		const result = run(source, 'Public Sub Build', { Helpers: '' });
		if (!result.ok) { throw new Error(result.reason); }
		expect(applyVbaTextEdits(source, result.edits)).toContain('    Helpers.Build\r\n');
	});

	it('leaves an unqualified call exactly as it is', () => {
		const caller = 'Public Sub Go()\r\n    Build\r\nEnd Sub\r\n';
		const result = run(SOURCE, 'Public Sub Build', { Helpers: '', Caller: caller });
		if (!result.ok) { throw new Error(result.reason); }
		expect(result.otherModules!.some((m) => m.moduleName === 'Caller')).toBe(false);
	});

	it('leaves a qualified name inside a comment or a string alone', () => {
		const caller = [
			'Public Sub Go()',
			"    ' Reports.Build does the work",
			'    Debug.Print "Reports.Build"',
			'End Sub',
			'',
		].join('\r\n');
		const result = run(SOURCE, 'Public Sub Build', { Helpers: '', Caller: caller });
		if (!result.ok) { throw new Error(result.reason); }
		expect(result.otherModules!.some((m) => m.moduleName === 'Caller')).toBe(false);
	});

	it('titles the action with both module names', () => {
		const result = run(SOURCE, 'Public Sub Build', { Helpers: '' });
		expect(result.ok && result.title).toBe("Move 'Build' to Helpers");
	});
});

describe('what it refuses, and why', () => {
	it('names the Private variable that would be stranded', () => {
		const source = [
			'Private Cache As Long',
			'',
			'Public Sub Build()',
			'    Cache = 1',
			'End Sub',
			'',
		].join('\r\n');
		expect(reason(source, 'Public Sub Build', { Helpers: '' }))
			.toMatch(/uses 'Cache', which is Private to Reports/);
	});

	it('names a Private procedure it calls', () => {
		const source = [
			'Public Sub Build()',
			'    Helper',
			'End Sub',
			'',
			'Private Sub Helper()',
			'End Sub',
			'',
		].join('\r\n');
		expect(reason(source, 'Public Sub Build', { Helpers: '' })).toMatch(/uses 'Helper'/);
	});

	it('treats a bare module-level Dim as Private, because VBA does', () => {
		const source = 'Dim Cache As Long\r\n\r\nPublic Sub Build()\r\n    Cache = 1\r\nEnd Sub\r\n';
		expect(reason(source, 'Public Sub Build', { Helpers: '' })).toMatch(/'Cache'/);
	});

	it('does not strand a Public module member', () => {
		const source = 'Public Shared As Long\r\n\r\nPublic Sub Build()\r\n    Shared = 1\r\nEnd Sub\r\n';
		const result = run(source, 'Public Sub Build', { Helpers: '' });
		expect(result.ok).toBe(true);
	});

	it('ignores a Private name that only appears in a comment', () => {
		const source = [
			'Private Cache As Long',
			'',
			'Public Sub Build()',
			"    ' Cache is not touched here",
			'End Sub',
			'',
		].join('\r\n');
		expect(run(source, 'Public Sub Build', { Helpers: '' }).ok).toBe(true);
	});

	it('refuses a target the project has not got', () => {
		expect(reason(SOURCE, 'Public Sub Build', { Helpers: '' }, 'Missing'))
			.toMatch(/no module called 'Missing'/);
	});

	it('refuses a target that already has the name', () => {
		const helpers = 'Public Sub Build()\r\nEnd Sub\r\n';
		expect(reason(SOURCE, 'Public Sub Build', { Helpers: helpers }))
			.toMatch(/already has a procedure called 'Build'/);
	});

	it('refuses moving a procedure to where it already is', () => {
		expect(reason(SOURCE, 'Public Sub Build', { Helpers: '' }, 'Reports'))
			.toMatch(/already in Reports/);
	});

	it('refuses a caret that is not in a procedure', () => {
		expect(reason(SOURCE, 'Option Explicit', { Helpers: '' })).toMatch(/caret in the procedure/);
	});
});
