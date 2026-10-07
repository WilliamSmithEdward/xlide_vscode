import { describe, expect, it, vi } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { localUsesIn } from '../src/analyzer/refactor/shared';
import { findIdentifierOccurrences } from '../src/vbaSourceScan';

const wordPattern = '[\\p{L}_][\\p{L}\\p{M}\\p{N}_]*';

describe('procedure-local refactor occurrence work', () => {
	it.each([10, 1000].flatMap(count => ['\n', '\r\n', '\r'].map(eol => ({count, eol}))))('avoids scanning $count unrelated procedures with EOL $eol', ({count, eol}) => {
		const before = Array.from({length: count}, (_, i) => `Sub Other${i}()${eol}outsideToken = 1${eol}End Sub${eol}`).join('');
		const source = before + ['Sub Main()', 'Dim target As Long', 'target = 7', 'Debug.Print target', 'End Sub', ''].join(eol);
		const module = parseModule(source);
		const procedure = module.members.find(member => member.kind === 'Procedure' && member.name === 'Main');
		if (!procedure || procedure.kind !== 'Procedure') throw new Error('missing procedure');
		const declaration = procedure.body.find(node => node.kind === 'VariableGroup');
		if (!declaration) throw new Error('missing declaration');
		findIdentifierOccurrences(source, 'prime'); // Stripping cache belongs to the source, not this query.
		let unrelatedScans = 0;
		const original = RegExp.prototype.exec;
		const spy = vi.spyOn(RegExp.prototype, 'exec').mockImplementation(function(this: RegExp, input: string) {
			if (this.source === wordPattern && this.flags === 'gu' && input.includes('outsideToken')) unrelatedScans++;
			return original.call(this, input);
		});
		let result;
		try { result = localUsesIn(source, procedure, declaration.span, 'TARGET'); } finally { spy.mockRestore(); }
		const write = {line: count * 3 + 2, column: 0, offset: before.length + 'Sub Main()'.length + eol.length + 'Dim target As Long'.length + eol.length, text: 'target'};
		const read = {line: count * 3 + 3, column: 12, offset: write.offset + 'target = 7'.length + eol.length + 12, text: 'target'};
		expect(result).toEqual({uses: [write, read], writes: [write]});
		expect(unrelatedScans).toBe(0);
	});
});
