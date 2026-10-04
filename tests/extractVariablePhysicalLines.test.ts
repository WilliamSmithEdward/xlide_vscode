import { expect, it } from 'vitest';
import { extractVariable } from '../src/analyzer/refactor/extractVariable';
import { encapsulateField } from '../src/analyzer/refactor/encapsulateField';
import { implementInterface } from '../src/analyzer/refactor/implementInterface';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { detectEol } from '../src/vbaSourceScan';

for(const [source,eol] of [['','\n'],['single line','\n'],['a\rb','\r'],['a\nb','\n'],['a\r\nb','\r\n'],['a\rb\nc','\n'],['a\r\nb\rc','\r\n'],['a\nb\r\nc','\r\n']])it('detects serialization EOL '+JSON.stringify(source),()=>expect(detectEol(source)).toBe(eol));
for(const eol of ['\n','\r\n','\r'])for(const nested of [false,true])for(const lines of [0,1000])it('extracts on its physical line '+JSON.stringify({eol,nested,lines}),()=>{
 const prefix=Array(lines).fill("' earlier module text").join(eol)+(lines?eol:'')+'Sub Go()'+eol+(nested?'    If True Then'+eol:'');
 const indent=nested?'\t        ':'    ',suffix=(nested?'    End If'+eol:'')+'End Sub'+eol;
 const source=prefix+indent+'Debug.Print 2 * 3'+eol+suffix,start=source.indexOf('2 * 3');
 const result=extractVariable({source,span:{start,end:start+5}});if(!result.ok)throw Error(result.reason);
 const expected=prefix+indent+'Dim value As Double'+eol+indent+'value = 2 * 3'+eol+indent+'Debug.Print value'+eol+suffix;
 expect(applyVbaTextEdits(source,result.edits)).toBe(expected);expect(result.edits[0].span).toEqual({start:prefix.length,end:prefix.length});expect(expected.slice(result.renameSpan!.start,result.renameSpan!.end)).toBe('value');
});
for(const headerEol of ['\n','\r\n'])it('uses nearest CR break inside a mixed module '+JSON.stringify(headerEol),()=>{
 const prefix='Sub Go()'+headerEol+'    Debug.Print 1\r',source=prefix+'    Debug.Print 2 * 3\rEnd Sub'+headerEol,start=source.indexOf('2 * 3');
 const result=extractVariable({source,span:{start,end:start+5}});if(!result.ok)throw Error(result.reason);
 expect(applyVbaTextEdits(source,result.edits)).toBe(prefix+'    Dim value As Double'+headerEol+'    value = 2 * 3'+headerEol+'    Debug.Print value\rEnd Sub'+headerEol);
});
for(const eol of ['\n','\r\n','\r']){
 it('preserves interface destination terminators '+JSON.stringify(eol),()=>{
  const source='Implements IJob'+eol,result=implementInterface({source,moduleSources:{IJob:'Public Sub Work()'+eol+'End Sub'+eol}});if(!result.ok)throw Error(result.reason);
  expect(applyVbaTextEdits(source,result.edits)).toBe(source+eol+'Private Sub IJob_Work()'+eol+"    Err.Raise 5 'TODO: implement this interface member"+eol+'End Sub'+eol);
 });
 it('preserves field-generated property terminators '+JSON.stringify(eol),()=>{
  const source='Public Total As Long'+eol,result=encapsulateField({source,offset:source.indexOf('Total')});if(!result.ok)throw Error(result.reason);
  expect(applyVbaTextEdits(source,result.edits)).toBe(['Private m_Total As Long','','Public Property Get Total() As Long','    Total = m_Total','End Property','','Public Property Let Total(ByVal RHS As Long)','    m_Total = RHS','End Property',''].join(eol));
 });
}
