// Issue #1418: the expression work of a full analysis grows with the module,
// never faster. Each statement shape is analyzed at 100, 200 and 400
// statements, with every statement's text unique so no cache hides the work,
// and its parses and fresh tokenizations must be affine in the size.
import { describe, expect, it, vi } from 'vitest';

const work = { parse: 0, tokenize: 0 };
vi.mock('../../../src/analyzer/parser/parseExpression', async (load) => {
	const real = await load<typeof import('../../../src/analyzer/parser/parseExpression')>();
	return {
		...real,
		parseExpression: (...args: Parameters<typeof real.parseExpression>) => {
			work.parse++;
			return real.parseExpression(...args);
		},
	};
});
vi.mock('../../../src/analyzer/lexer/tokenize', async (load) => {
	const real = await load<typeof import('../../../src/analyzer/lexer/tokenize')>();
	return {
		...real,
		tokenize: (...args: Parameters<typeof real.tokenize>) => {
			work.tokenize++;
			return real.tokenize(...args);
		},
	};
});

const { analyzeVbaModuleSource } = await import('../../../src/vbaModuleAnalysis');

/** Each shape takes the statement's index, so no two statements share text. */
const SHAPES: Record<string, (i: number) => string[]> = {
	assignment: (i) => [`    n = n + ${i}`],
	call: (i) => [`    Foo n, ${i}`],
	blockIf: (i) => [`    If n > ${i} Then`, '        n = 2', `    ElseIf n < ${i} Then`, '        n = 3', '    End If'],
	inlineIf: (i) => [`    If n > ${i} Then n = 2 Else If o Is Nothing Then n = ${i}`],
	print: (i) => [`    Debug.Print n; Tab(${i}); s`],
	select: (i) => ['    Select Case n', `    Case ${i}, Is > ${i + 1}`, '        n = 1', '    End Select'],
	forLoop: (i) => [`    For i = ${i} To n Step 1`, '    Next i'],
	doLoop: (i) => [`    Do While n > ${i}`, '        n = n - 1', '    Loop'],
	isOperand: (i) => [`    If o Is Nothing Then n = CDbl(${i})`],
};

function workFor(shape: (i: number) => string[], count: number, salt: string): { parse: number; tokenize: number } {
	const body: string[] = [];
	for (let p = 0; p < count / 20; p++) {
		body.push(`Sub P${p}${salt}()`, '    Dim n As Long, i As Long, s As String, o As Object');
		for (let k = 0; k < 20; k++) {
			body.push(...shape(p * 20 + k));
		}
		body.push('End Sub');
	}
	const source = ['Option Explicit', `Sub Foo(a, b) ' ${salt}`, 'End Sub', ...body, ''].join('\r\n');
	work.parse = 0;
	work.tokenize = 0;
	analyzeVbaModuleSource({ source, moduleName: 'Module1', moduleKind: 'standard', host: 'excel' });
	return { ...work };
}

describe('expression work grows linearly with the module (#1418)', () => {
	it.each(Object.entries(SHAPES))('%s', (name, shape) => {
		const [small, medium, large] = [100, 200, 400].map((count, i) => workFor(shape, count, `${name}${i}`));
		// Work that is a fixed amount per statement plus a constant: doubling
		// the module from 200 to 400 adds exactly twice what 100 to 200 did.
		expect(large.parse - medium.parse).toBe(2 * (medium.parse - small.parse));
		expect(large.tokenize - medium.tokenize).toBe(2 * (medium.tokenize - small.tokenize));
		expect(medium.parse).toBeGreaterThan(small.parse);
	});
});
