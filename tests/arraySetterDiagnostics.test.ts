import { describe, expect, it } from 'vitest';
import { analyzeProjectModule } from './diagnostics/helpers';
import { resolveAssignmentValueCompletion } from '../src/analyzer/completion/assignmentValueCompletion';
import { projectOptions } from './diagnostics/helpers';
const setter = (indexed: boolean) => `Public Property Let Flags(${indexed ? 'ByVal index As Long, ' : ''}ByRef value() As Boolean)\nEnd Property`;
function diagnostics(statement: string, indexed: boolean, bare = false) {
 const source = `Option Explicit\n${bare ? setter(indexed) : ''}\nSub T(ByVal item As Widget)\nDim values(1) As Boolean\n${statement}\nEnd Sub`;
 return analyzeProjectModule(source, bare ? [] : [{ moduleName: 'Widget', moduleKind: 'class', source: setter(indexed) }], 'Caller').filter(d => d.severity === 'error');
}
describe('Property Let array value contracts', () => {
 for (const indexed of [false, true]) {
  const target = `Flags${indexed ? '(1)' : ''}`;
  for (const bare of [true, false]) {
   const name = bare ? target : `item.${target}`;
   it(`reports scalar shape once for ${name}`, () => {
    for (const rhs of ['True', '5', '"nonsense"', 'values(1)']) {
     const found = diagnostics(`${name} = ${rhs}`, indexed, bare);
     expect(found, `${name} = ${rhs}`).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
     expect(found[0].message).toContain('VBE compile error');
    }
   });
   it(`reports a whole array assignment once for ${name}`, () => {
    expect(diagnostics(`${name} = values`, indexed, bare)).toEqual([expect.objectContaining({ code: 'array-target-assignment' })]);
   });
  }
  it(`checks With and single-line If, indexed=${indexed}`, () => {
   expect(diagnostics(`With item\n.${target} = True\nEnd With`, indexed)).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
   expect(diagnostics(`If True Then item.${target} = True`, indexed)).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
  });
 }
 it('uses an exported array setter contract', () => {
  const source = 'Option Explicit\nSub T()\nFlags = True\nEnd Sub';
  expect(analyzeProjectModule(source, [{ moduleName: 'Library', source: setter(false) }], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
 });
 it('preserves scalar setters and array elements', () => {
  const source = 'Option Explicit\nProperty Let Flags(ByVal value As Boolean)\nEnd Property\nSub T()\nDim values(1) As Boolean\nFlags = values(1)\nEnd Sub';
  expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
 });
 it('keeps a local scalar shadow of an exported setter', () => {
  const source = 'Option Explicit\nSub T()\nDim Flags As Boolean\nFlags = True\nDebug.Print Flags\nEnd Sub';
  expect(analyzeProjectModule(source, [{ moduleName: 'Library', source: setter(false) }], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
 });
 it('does not suggest scalar constants for an array value', () => {
  const source = 'Sub T(ByVal item As Widget)\nitem.Flags(1) = ';
  const options = projectOptions([{ moduleName: 'Caller', source }, { moduleName: 'Widget', moduleKind: 'class', source: setter(true) }], 'Caller');
  expect(resolveAssignmentValueCompletion(source, source.length, { projectClassMembers: options.projectClassMembers })).toBeUndefined();
 });
});

it.each(['values()', '(values)', '((values))'])('reports a whole array spelling %s', rhs => {
 expect(diagnostics('Flags = ' + rhs, false, true)).toEqual([expect.objectContaining({ code: 'array-target-assignment' })]);
});
it('checks the declared Variant shape despite holding an array', () => {
 expect(diagnostics('Dim value As Variant\nvalue = Array(True)\nFlags = value', false, true)).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
});
it('rejects the untyped runtime Array result by shape', () => {
 expect(diagnostics('Flags = Array(True)', false, true)).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
});
it('keeps the native default-member error for Collection values', () => {
 expect(diagnostics('Flags = New Collection', false, true)).toEqual([expect.objectContaining({ code: 'argument-count' })]);
});
it('rejects an object without a default value by shape', () => {
 const source = 'Option Explicit\n' + setter(false) + '\nSub T()\nFlags = New Widget\nEnd Sub';
 expect(analyzeProjectModule(source, [{ moduleName: 'Widget', moduleKind: 'class', source: 'Public Flag As Boolean' }], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
});
it.each(['MakeFlags()', 'Split(True)'])('accepts a source function returning a typed array: %s', rhs => {
 const fn = rhs.startsWith('Split') ? 'Split(ByVal ignored As Boolean)' : 'MakeFlags()';
 const name = rhs.startsWith('Split') ? 'Split' : 'MakeFlags';
 const source = 'Option Explicit\n' + setter(false) + `\nFunction ${fn} As Boolean()\nDim result(1) As Boolean\n${name} = result\nEnd Function\nSub T()\nFlags = ${rhs}\nEnd Sub`;
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
});

it('keeps qualified VBA.Split intrinsic beside a source shadow', () => {
 const source = 'Option Explicit\n' + setter(false) + '\nFunction Split(ByVal ignored As Boolean) As Boolean()\nDim result(1) As Boolean\nSplit = result\nEnd Function\nSub T()\nFlags = VBA.Split("x")\nEnd Sub';
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
});

it('rejects a typed function with the wrong array element type', () => {
 const source = 'Option Explicit\n' + setter(false) + '\nFunction MakeFlags() As Long()\nDim result(1) As Long\nMakeFlags = result\nEnd Function\nSub T()\nFlags = MakeFlags()\nEnd Sub';
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'argument-shape-mismatch' })]);
});
it('preserves String conversion to Byte-array setter values', () => {
 const source = 'Option Explicit\nProperty Let Flags(ByRef bytes() As Byte)\nEnd Property\nSub T()\nFlags = "text"\nEnd Sub';
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
});
it('does not confuse a whole String array with a convertible String value', () => {
 const source = 'Option Explicit\nProperty Let Flags(ByRef bytes() As Byte)\nEnd Property\nSub T()\nDim values(1) As String\nFlags = values\nEnd Sub';
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'array-target-assignment' })]);
});
it('uses an exported setter owning DefByte for String conversion', () => {
 const source = 'Option Explicit\nSub T()\nFlags = "text"\nEnd Sub';
 expect(analyzeProjectModule(source, [{ moduleName: 'Library', source: 'DefByte B\nPublic Property Let Flags(ByRef bytes())\nEnd Property' }], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
});

it('reports the tracked Variant array as a runtime Byte-array mismatch', () => {
 const source = 'Option Explicit\nProperty Let Flags(ByRef bytes() As Byte)\nEnd Property\nSub T()\nDim value As Variant\nvalue = Array(1,2)\nFlags = value\nEnd Sub';
 const found = analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error');
 expect(found).toEqual([expect.objectContaining({ code: 'assignment-type-mismatch' })]);
 expect(found[0].message).toContain("Run-time error '13'");
});
it('accepts a Variant holding a String for a Byte-array setter', () => {
 const source = 'Option Explicit\nProperty Let Flags(ByRef bytes() As Byte)\nEnd Property\nSub T()\nDim value As Variant\nvalue = "text"\nFlags = value\nEnd Sub';
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
});
it('tracks the same Byte-array value through a With property write', () => {
 const source = 'Option Explicit\nSub T(ByVal item As Widget)\nDim value As Variant\nvalue = Array(1,2)\nWith item\n.Flags = value\nEnd With\nEnd Sub';
 expect(analyzeProjectModule(source, [{ moduleName: 'Widget', moduleKind: 'class', source: 'Property Let Flags(ByRef bytes() As Byte)\nEnd Property' }], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'assignment-type-mismatch' })]);
});
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
it('indexes local value facts once when many distinct Variants feed a scalar setter', () => {
 const count = 500;
 const source = 'Option Explicit\nProperty Let Flags(ByVal value As Boolean)\nEnd Property\nSub T()\n' + Array.from({length:count}, (_, i) => `Dim v${i} As Variant\nv${i} = Array(1)\nFlags = v${i}\n`).join('') + 'End Sub';
 const mod = parseModule(source), symbols = buildModuleSymbols('Caller', 'standard', source, { parsedModule: mod });
 let reads = 0;
 const proc = symbols.root.children!.find(s => s.name === 'T')!;
 for (const child of proc.children!) { const name = child.name; Object.defineProperty(child, 'name', { get() { reads++; return name; } }); }
 const findings: string[] = [];
 checkAssignmentTypes(source, mod, symbols, undefined, {}, undefined, (code) => findings.push(code));
 expect(findings).toHaveLength(count);
 expect(findings.every(code => code === 'assignmentTypeMismatch')).toBe(true);
 expect(reads).toBeLessThanOrEqual(count * 35);
});
it('accepts a With Range.Value array transfer and retains array comparisons', () => {
 const prelude = 'Option Explicit\nSub T()\nDim values As Variant\nvalues = Array(1,2)\n';
 expect(analyzeProjectModule(prelude + 'With ThisWorkbook.Worksheets(1).Range("A1:B1")\n.Value = values\nEnd With\nEnd Sub', [], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
 expect(analyzeProjectModule(prelude + 'If values = 1 Then Exit Sub\nEnd Sub', [], 'Caller')).toContainEqual(expect.objectContaining({code:'variant-value-misuse'}));
});
it('keeps tracked array compatibility errors for scalar setters', () => {
 const source = 'Option Explicit\nProperty Let Flags(ByVal value As Boolean)\nEnd Property\nSub T()\nDim value As Variant\nvalue = Array(1,2)\nFlags = value\nEnd Sub';
 expect(analyzeProjectModule(source, [], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({code:'assignment-type-mismatch'})]);
});

it('checks typed host setters while permitting the native Variant alignment assignment', () => {
 const prelude = 'Option Explicit\nSub T(ByVal ws As Worksheet)\nDim values As Variant\nvalues = Array(1,2)\n';
 expect(analyzeProjectModule(prelude+'ws.EnableCalculation = values\nEnd Sub', [], 'Caller').filter(d => d.severity === 'error')).toEqual([expect.objectContaining({code:'assignment-type-mismatch'})]);
 expect(analyzeProjectModule(prelude+'ws.Range("A1").HorizontalAlignment = values\nEnd Sub', [], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
});
