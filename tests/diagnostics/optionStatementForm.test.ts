import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

const CODE = 'invalid-option-statement';

const TAIL = '\nPublic Sub Recalculate()\n    Dim n As Long\n    n = 1\nEnd Sub\n';

function hits(optionLine: string, opts: Record<string, unknown> = {}) {
	const src = `${optionLine}${TAIL}`;
	return { src, found: byCode(analyzeModule(src, opts as never), CODE) };
}

function only(optionLine: string, opts: Record<string, unknown> = {}) {
	const { src, found } = hits(optionLine, opts);
	expect(found, optionLine).toHaveLength(1);
	return { message: found[0].message, at: spanText(src, found[0]) };
}

describe('malformed Option statements (#74)', () => {
	// Each expectation below mirrors what the live VBE said when the same line
	// was compiled: the Excel oracle probes recorded in the rule's comment.

	it('reports an argument list after Option Explicit, at the parenthesis', () => {
		const { message, at } = only('Option Explicit()');

		expect(at).toBe('(');
		expect(message).toContain('Option Explicit');
		expect(message).toContain('end');
	});

	it('reports a trailing word after Option Explicit', () => {
		expect(only('Option Explicit Foo').at).toBe('Foo');
	});

	it('reports a word that names no directive', () => {
		const { message, at } = only('Option Nonsense');

		expect(at).toBe('Nonsense');
		expect(message).toContain('Base, Compare, Explicit or Private');
	});

	it('reports Option with no directive at all', () => {
		const { message, at } = only('Option');

		expect(at).toBe('Option');
		expect(message).toContain('Base, Compare, Explicit or Private');
	});

	it('reports an Option Base argument that is not 0 or 1', () => {
		const { message, at } = only('Option Base 2');

		expect(at).toBe('2');
		expect(message).toContain('0 or 1');
	});

	it('reports Option Base with no argument', () => {
		expect(only('Option Base').at).toBe('Base');
	});

	it('reports an unknown Option Compare argument', () => {
		const { message, at } = only('Option Compare Sideways');

		expect(at).toBe('Sideways');
		expect(message).toContain('Binary or Text');
	});

	it('reports Option Private without Module', () => {
		const { message, at } = only('Option Private');

		expect(at).toBe('Private');
		expect(message).toContain('Option Private Module');
	});

	it('reports trailing text after an otherwise valid directive', () => {
		expect(only('Option Base 1 Extra').at).toBe('Extra');
		expect(only('Option Compare Text Extra').at).toBe('Extra');
		expect(only('Option Private Module Extra').at).toBe('Extra');
	});
});

describe('well-formed Option statements stay quiet', () => {
	for (const line of [
		'Option Explicit',
		'Option Base 0',
		'Option Base 1',
		'Option Compare Binary',
		'Option Compare Text',
		'Option Private Module',
		'option explicit',
		'OPTION BASE 1',
	]) {
		it(`accepts ${line}`, () => {
			expect(hits(line).found).toHaveLength(0);
		});
	}

	it('accepts a trailing comment', () => {
		expect(hits("Option Explicit ' why we are strict").found).toHaveLength(0);
	});

	it('accepts a line continuation inside the statement', () => {
		expect(hits('Option _\n    Explicit').found).toHaveLength(0);
	});

	it('ignores an Option in an inactive conditional arm', () => {
		const src = '#Const READY = 0\n#If READY Then\nOption Explicit()\n#End If\n' + TAIL;

		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('Option Compare Database follows the host', () => {
	it('is accepted in an Access project', () => {
		expect(hits('Option Compare Database', { host: 'access' }).found).toHaveLength(0);
	});

	it('is accepted when no project names a host', () => {
		// A loose file says nothing about its host, so the rule says nothing.
		expect(hits('Option Compare Database').found).toHaveLength(0);
	});

	for (const host of ['excel', 'word', 'powerpoint']) {
		it(`is reported in a ${host} project`, () => {
			const { message, at } = only('Option Compare Database', { host });

			expect(at).toBe('Database');
			expect(message).toContain('Access');
		});
	}
});
