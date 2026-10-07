import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Runtime assignment regressions', () => {
 suiteSetup(activate);
 teardown(closeAllEditors);
 test('readonly Err diagnostics publish and clear after an actual edit', async () => {
  const document = await open(await writeModule('ReadonlyErrProbe', 'Option Explicit\r\nFunction Main() As Variant\r\nErr.LastDllError = 5\r\nMain = 1\r\nEnd Function\r\n'));
  const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'readonly-member-assignment'), 'readonly Err assignment should be reported');
  assert.equal(document.getText(finding.range), 'LastDllError');
  assert.match(finding.message, /Can't assign to read-only property/);
  await vscode.window.activeTextEditor!.edit(edit => edit.replace(new vscode.Range(2, 0, 2, document.lineAt(2).text.length), 'Err.Number = 5'));
  await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.code === 'readonly-member-assignment'), 'valid writable property should clear the finding');
  await document.save();
 });
 for (const rhs of ['values', 'MakeFlags()']) {
  test(`host scalar setter rejects ${rhs} and accepts an array element`, async () => {
   const line = 'ThisWorkbook.Worksheets(1).EnableCalculation = ';
   const source = ['Option Explicit', 'Function Main() As Variant', 'Dim values(1) As Boolean', line + rhs, 'Main = 1', 'End Function', 'Function MakeFlags() As Boolean()', 'Dim flags(1) As Boolean', 'MakeFlags = flags', 'End Function', ''].join('\r\n');
   const document = await open(await writeModule(rhs === 'values' ? 'HostArrayValueProbe' : 'HostArrayFunctionProbe', source));
   const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'assignment-type-mismatch'), 'whole array assigned to a host Boolean property should be reported');
   assert.match(finding.message, /Run-time error '13'/);
   assert.doesNotMatch(finding.message, /compile error/);
   await vscode.window.activeTextEditor!.edit(edit => edit.replace(new vscode.Range(3, 0, 3, document.lineAt(3).text.length), line + 'values(0)'));
   await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.code === 'assignment-type-mismatch'), 'array element assigned to the property should clear the finding');
   await document.save();
  });
 }
});
