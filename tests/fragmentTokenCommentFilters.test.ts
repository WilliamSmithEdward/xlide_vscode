import { expect, it } from 'vitest';
import { sourceExpressionSyntaxProblem } from '../src/analyzer/diagnostics/rules/shared';
import { constantStringValue } from '../src/analyzer/diagnostics/typeInference';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { rawExpressionTokens } from '../src/analyzer/diagnostics/walker';
function constant(expression: string) {
 return buildModuleSymbols('M', 'standard', 'Const k As String = ' + expression).root.children!.find(s => s.kind === 'constant')!;
}
for (const scope of ['source-expression', 'constant-string']) for (const n of [10, 100, 1000]) it(`does no redundant comment filtering for ${scope} with ${n} values`, () => {
 const text = scope === 'source-expression' ? 'Array(' + Array(n).fill('1').join(', ') + ')' : Array(n).fill('"x"').join(' & ');
 const symbol = scope === 'constant-string' ? constant(text) : undefined;
 const original = Array.prototype.filter;
 let redundantElements = 0;
 try {
  Array.prototype.filter = function (callback: any, thisArg?: any): any {
   const stack = new Error().stack ?? '';
   if (/\.kind\s*!==\s*['"]comment/.test(String(callback)) && !/\bstatementTokens(?:Cached)?\b/.test(stack)
     && stack.includes(scope === 'source-expression' ? 'sourceExpressionSyntaxProblem' : 'constantStringValue')) redundantElements += this.length;
   return original.call(this, callback, thisArg);
  };
  expect(scope === 'source-expression' ? sourceExpressionSyntaxProblem(text) : constantStringValue(symbol!)).toBeUndefined();
 } finally { Array.prototype.filter = original; }
 expect(redundantElements).toBe(0);
});
for (const [text, expected] of [
 ['Array(1, 2)', undefined], ['Array(1, _\n2)\n\' comment', undefined], ['[A1:B2]', undefined],
 ['5 \' comment', 'the literal 5'], ['"can\'t"', 'the literal "can\'t"'], ['#1/1/2000#', 'the literal #1/1/2000#'],
 ['(c)', "an expression that opens with '('"], ['\' comment\n', undefined],
] as const) it(`preserves source-expression syntax for ${JSON.stringify(text)}`, () => expect(sourceExpressionSyntaxProblem(text)).toBe(expected));
for (const [text, expected] of [['"x" \' comment','x'], ['"can\'t"',"can't"], ['"a" & "b"',undefined], ['#1/1/2000#',undefined]] as const)
 it(`preserves constant string classification for ${JSON.stringify(text)}`, () => expect(constantStringValue(constant(text))).toBe(expected));
for (const text of ['#1/1/2000# \' comment', 'Array(1, _\n2) \' comment', '"can\'t"\n\' comment'])
 it(`fragment contract excludes trivia for ${JSON.stringify(text)}`, () => expect(rawExpressionTokens(text).some(t => t.kind === 'comment' || t.kind === 'newline')).toBe(false));
