import { expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { setCallEffects, setDeclaredFacts, straightLineAssignments, straightLineUnreachable, type ReachingAssignments } from '../src/analyzer/diagnostics/straightLineValues';
import { rawExpressionTokens } from '../src/analyzer/diagnostics/walker';
import type { ProcedureNode } from '../src/analyzer/parser/nodes';

it('keeps equal-valued starts with different ByRef effects separate', () => {
	const source = 'Sub P()\nChange n\nDebug.Print n\nEnd Sub\n';
	const proc = parseModule(source).members[0] as ProcedureNode;
	for (const value of [0, 2, 0]) {
		const initial: ReachingAssignments = new Map([['n', rawExpressionTokens('1')]]);
		setCallEffects(initial, tokens => tokens[0]?.rawText.toLowerCase() === 'change' ? new Map([['n', rawExpressionTokens(String(value))]]) : new Map());
		const result = straightLineAssignments(source, proc.body, undefined, initial);
		expect(result.get(proc.body[1])?.get('n')?.map(token => token.rawText).join('')).toBe(String(value));
	}
});

it('updates a retained start after call effects change', () => {
	const source = 'Sub P()\nChange n\nDebug.Print n\nEnd Sub\n';
	const proc = parseModule(source).members[0] as ProcedureNode;
	const initial: ReachingAssignments = new Map([['n', rawExpressionTokens('1')]]);
	for (const value of [0, 2, 0]) {
		setCallEffects(initial, tokens => tokens[0]?.rawText.toLowerCase() === 'change' ? new Map([['n', rawExpressionTokens(String(value))]]) : new Map());
		expect(straightLineAssignments(source, proc.body, undefined, initial).get(proc.body[1])?.get('n')?.[0].rawText).toBe(String(value));
	}
});

it('does not reuse reachability when declaration facts change', () => {
	const source = 'Sub P()\nIf Flag = 0 Then\nDebug.Print 1\nElse\nDebug.Print 2\nEnd If\nEnd Sub\n';
	const proc = parseModule(source).members[0] as ProcedureNode;
	const initial: ReachingAssignments = new Map();
	const branch = proc.body[0];
	if (branch.kind !== 'IfBlock') throw new Error('fixture');
	for (const value of [0, 1, 0]) {
		setDeclaredFacts(initial, { type: () => undefined, bounds: () => undefined, constant: lower => lower === 'flag' ? value : undefined });
		const dead = straightLineUnreachable(source, proc.body, undefined, initial);
		expect(dead.has(branch.branches[0].body[0])).toBe(value !== 0);
		expect(dead.has(branch.branches[1].body[0])).toBe(value === 0);
	}
});
it('shares a retained start across value, exit and reachability queries', () => {
	const source = 'Sub P()\nn = 2\nDebug.Print n\nEnd Sub\n';
	const proc = parseModule(source).members[0] as ProcedureNode;
	const initial: ReachingAssignments = new Map([['n', rawExpressionTokens('0')]]);
	const first = straightLineAssignments(source, proc.body, undefined, initial);
	for (let i = 0; i < 10; i++) {
		expect(straightLineAssignments(source, proc.body, undefined, initial)).toBe(first);
		expect(straightLineUnreachable(source, proc.body, undefined, initial).size).toBe(0);
	}
});
