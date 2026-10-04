import { expect, it } from 'vitest';
import { extractVariable } from '../src/analyzer/refactor/extractVariable';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

const cases: [string, string][] = [
 ['Dim first As Long, value As Long', 'value2'],
 ['Static first As Long, value As Long', 'value2'],
 ['Const first As Long = 1, value As Long = 2', 'value2'],
 ['Dim [value] As Long', 'value2'],
 ['Dim first As Long, VALUE As Long, value2 As Long', 'value3'],
 ['Dim first As Long, _\n    [value] As Long', 'value2'],
 ['If True Then\n    Dim first As Long, [value] As Long\n    End If', 'value2'],
 ['ReDim value(3)', 'value2'],
 ['ReDim Preserve first(1, 2), [value](3)', 'value2'],
 ['10 ReDim Preserve first(1, 2), value(3)', 'value2'],
 ['If True Then ReDim Preserve first(1, 2), value(3)', 'value2'],
 ['If True Then Debug.Print 1 Else ReDim value(3)', 'value2'],
 ['Debug.Print "Dim value As Long"', 'value'],
 ["' Dim value As Long", 'value'],
 ['Rem Dim value As Long', 'value'],
 ['Const first As String = "Dim value As Long"', 'value'],
 ['ReDim item.value(3)', 'value'],
 ['Dim first As Long, value2 As Long', 'value'],
];
for(const eol of ['\n','\r\n'])for(const [statement,name] of cases){
 it('reserves actual declarations: '+JSON.stringify(statement)+' '+JSON.stringify(eol),()=>{
  const prefix=['Sub Go()','    '+statement.replace(/\n/g,eol),''].join(eol);
  const source=prefix+'    Debug.Print 2 * 3'+eol+'End Sub'+eol,start=source.indexOf('2 * 3');
  const result=extractVariable({source,span:{start,end:start+5}});if(!result.ok)throw Error(result.reason);
  expect(result.title).toBe("Extract '"+name+"'");
  expect(applyVbaTextEdits(source,result.edits)).toBe(prefix+'    Dim '+name+' As Double'+eol+'    '+name+' = 2 * 3'+eol+'    Debug.Print '+name+eol+'End Sub'+eol);
  const applied=applyVbaTextEdits(source,result.edits);expect(applied.slice(result.renameSpan!.start,result.renameSpan!.end)).toBe(name);
 });
}
