import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { projectClassMembersIndex } from '../src/analyzer/completion/memberAccess';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
import { projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';

// The project's member surfaces are indexed by name once per list. These pin
// that a name two surfaces share stays out, as both callers treated it, and
// that a new list, as every project edit makes, is indexed afresh.

function surface(name: string, kind: VbaProjectClassMembers['kind'], moduleName = name): VbaProjectClassMembers {
	return { name, kind, moduleName, members: [] };
}

describe('projectClassMembersIndex', () => {
	it('indexes each name by lower case and leaves out a name two surfaces share', () => {
		const ticket = surface('Ticket', 'class');
		const lib = surface('Lib', 'standardModule');
		const index = projectClassMembersIndex([
			ticket,
			lib,
			surface('Shared', 'userType', 'Lib'),
			surface('SHARED', 'class'),
			surface('shared', 'enum', 'Ticket'),
		]);
		expect(index.get('ticket')).toBe(ticket);
		expect(index.get('lib')).toBe(lib);
		expect(index.has('shared')).toBe(false);
		expect([...index.keys()]).toEqual(['ticket', 'lib']);
	});

	it('answers the same list from one index, and a new list from a new one', () => {
		const list = [surface('Ticket', 'class')];
		expect(projectClassMembersIndex(list)).toBe(projectClassMembersIndex(list));

		const edited = [surface('Ticket', 'class'), surface('Order', 'class')];
		expect(projectClassMembersIndex(edited)).not.toBe(projectClassMembersIndex(list));
		expect(projectClassMembersIndex(edited).has('order')).toBe(true);
		expect(projectClassMembersIndex(list).has('order')).toBe(false);
	});

	it('reaches member checks after an edit', () => {
		const caller = 'Option Explicit\r\nSub Main()\r\n    Lib.Known\r\n    Lib.Added\r\nEnd Sub\r\n';
		const index = new ProjectIndex();
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: caller });
		const missing = (): string[] => analyzeModule(caller, {
			...projectAnalysisOptionsForModule(index, 'Caller'),
			moduleName: 'Caller',
			moduleKind: 'standard',
		})
			.filter((d) => d.code === 'member-not-found')
			.map((d) => caller.slice(d.span.start, d.span.end));

		index.setModule({ moduleName: 'Lib', moduleKind: 'standard', source: 'Public Sub Known()\r\nEnd Sub\r\n' });
		expect(missing()).toEqual(['Added']);

		index.setModule({ moduleName: 'Lib', moduleKind: 'standard', source: 'Public Sub Known()\r\nEnd Sub\r\nPublic Sub Added()\r\nEnd Sub\r\n' });
		expect(missing()).toEqual([]);
	});
});
