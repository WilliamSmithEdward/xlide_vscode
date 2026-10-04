import { expect, it } from 'vitest';
import { checkClassInstanceValues } from '../src/analyzer/diagnostics/rules/classInstanceValues';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const classes: VbaProjectClassMembers[] = [{ name: 'Class1', moduleName: 'Class1', kind: 'class',
	members: [{ name: 'M', moduleName: 'Class1', kind: 'property', returns: 'Object', knownValue: 'nothing' }],
}];
function prepare(source: string) {
	const mod = parseModule(source);
	return { mod, symbols: buildModuleSymbols('M', 'standard', source, { parsedModule: mod }) };
}
function run(source: string) {
	const { mod, symbols } = prepare(source), output: unknown[][] = [];
	checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, undefined, (...args) => output.push(args));
	return output;
}

it.each([10, 100, 600])('bounds token-array iteration for %i indexed reads in one logical statement', count => {
	const terms = Array(count).fill('actor.M(1)'), lines: string[] = [];
	for (let i = 0; i < count; i += 30) { lines.push(terms.slice(i, i + 30).join('; ')); }
	const source = ['Sub Go()', 'Dim actor As New Class1', 'Debug.Print ' + lines.join('; _\n'), 'End Sub', ''].join('\n');
	const { mod, symbols } = prepare(source), output: unknown[][] = [];
	const original = Array.prototype[Symbol.iterator];
	let tokenReferences = 0;
	Array.prototype[Symbol.iterator] = function (this: unknown[]) {
		const first = this[0] as { rawText?: unknown; kind?: unknown } | undefined;
		if (typeof first?.rawText === 'string' && typeof first.kind === 'string') { tokenReferences += this.length; }
		return original.call(this);
	};
	try {
		checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, undefined, (...args) => output.push(args));
	} finally { Array.prototype[Symbol.iterator] = original; }
	expect(output).toHaveLength(count);
	let from = 0;
	for (const diagnostic of output) {
		const start = source.indexOf('actor.M(1)', from);
		expect(diagnostic).toEqual(['objectVariableNotSet',
			"'actor.M(1)' is Nothing here: nothing in Class1 sets M. This will raise Run-time error '91': Object variable or With block variable not set.",
			{ start, end: start + 'actor.M(1)'.length }]);
		from = start + 'actor.M(1)'.length;
	}
	expect(tokenReferences).toBeLessThanOrEqual(count * 10 + 20);
});

it.each(['\n', '\r\n', '\r'])('preserves nested argument spans across %j line endings', eol => {
	const source = ['Sub Go()', 'Dim actor As New Class1', 'Debug.Print actor.M((1 + 2))', 'End Sub', ''].join(eol);
	const start = source.indexOf('actor.M');
	expect(run(source)).toEqual([['objectVariableNotSet',
		"'actor.M((1+2))' is Nothing here: nothing in Class1 sets M. This will raise Run-time error '91': Object variable or With block variable not set.",
		{ start, end: start + 'actor.M((1 + 2))'.length }]]);
});

it('keeps unmatched argument lists quiet', () => {
	expect(run('Sub Go()\nDim actor As New Class1\nDebug.Print actor.M(1\nEnd Sub\n')).toEqual([]);
});
it('preserves indexed field-assignment suppression', () => {
	expect(run('Sub Go()\nDim actor As New Class1\nactor.M(1) = 2\nDebug.Print actor.M(1)\nEnd Sub\n')).toEqual([]);
});
