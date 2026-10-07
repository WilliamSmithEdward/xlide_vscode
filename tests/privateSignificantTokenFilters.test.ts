import { expect, it } from 'vitest';
import { conditionValue, numberValue } from '../src/analyzer/diagnostics/conditionValue';
import { statementLabelReferences, resolveProcedureLabelCompletions } from '../src/analyzer/flow/procedureLabels';
import { tokenize } from '../src/analyzer/lexer/tokenize';
const facts = { value: () => undefined };
function chunked(n: number, value: (i: number) => string, separator: string) {
 return Array.from({length: Math.ceil(n / 100)}, (_, i) => Array.from({length: Math.min(100, n - i * 100)}, (_, j) => value(i * 100 + j)).join(separator)).join(separator + ' _\n');
}
function counted<T>(target: string, run: () => T) {
 const original = Array.prototype.filter; let elements = 0; let result: T;
 try {
  Array.prototype.filter = function (callback: any, thisArg?: any): any {
   const stack = new Error().stack ?? '';
   if (/\.kind\s*!==\s*['"]comment/.test(String(callback)) && stack.includes(target) && !/\bstatementTokens(?:Cached)?\b/.test(stack)) elements += this.length;
   return original.call(this, callback, thisArg);
  };
  result = run();
 } finally { Array.prototype.filter = original; }
 return { result: result!, elements };
}
for (const n of [10, 100, 1000]) {
 it(`does no private comment filtering for ${n} numeric builtins`, () => {
  const tokens = tokenize(chunked(n, () => 'Abs(1)', ' + '));
  const actual = counted('ConditionParser', () => numberValue(tokens, facts));
  expect(actual.result).toBe(n); expect(actual.elements).toBe(0);
 });
 it(`does no private comment filtering for ${n} label references`, () => {
  const source = 'On index GoTo ' + chunked(n, i => 'L' + (i + 1), ', ');
  let from = source.indexOf('GoTo') + 4;
  const expected = Array.from({length: n}, (_, i) => { const text = 'L' + (i + 1), start = source.indexOf(text, from); from = start + text.length; return {key: 'name:' + text.toLowerCase(), text, span: {start, end: from}, kind: 'name', statementKind: 'on-goto'}; });
  const actual = counted('labelReferenceGroup', () => statementLabelReferences(source, {start: 0, end: source.length}));
  expect(actual.result).toEqual(expected); expect(actual.elements).toBe(0);
 });
 it(`does no private comment filtering for a ${n}-value completion prefix`, () => {
  const prefix = 'Sub Go()\nOn ' + chunked(n, () => '1', ' + ') + ' GoTo ';
  const source = prefix + '\nDone:\nEnd Sub\n';
  const actual = counted('isLabelTargetPrefix', () => resolveProcedureLabelCompletions(source, prefix.length));
  expect(actual.result).toEqual([{label: 'Done', kind: 'name', detail: 'Procedure label'}]); expect(actual.elements).toBe(0);
 });
}
for (const [text, expected] of [['Abs(-1)',1],['Round(2.5)',2],['Round(2.5, 0)',undefined],['IIf(True, 3, 4)',3],['Abs(-1) \' tail',1],['Round(1, 0, 3)',undefined]] as const)
 it(`preserves numeric builtin ${text}`, () => expect(numberValue(Object.freeze(tokenize(text)), facts)).toBe(expected));
for (const [text, expected] of [['Abs(-1) = 1',true],['IIf(False, 3, 4) = 3',false],['Len("can\'t") = 5 \' tail',true]] as const)
 it(`preserves condition ${text}`, () => expect(conditionValue(Object.freeze(tokenize(text)), facts)).toBe(expected));
for (const eol of ['\n','\r\n','\r']) it(`preserves named/numeric/commented references with ${JSON.stringify(eol)}`, () => {
 const source = '10 On x GoTo A, 20 \' tail' + eol;
 const a = source.indexOf('A'), line = source.indexOf('20');
 expect(statementLabelReferences(source, {start: 0, end: source.length})).toEqual([
  {key:'name:a',text:'A',span:{start:a,end:a+1},kind:'name',statementKind:'on-goto'},
  {key:'line:20',text:'20',span:{start:line,end:line+2},kind:'line',statementKind:'on-goto'},
 ]);
});
