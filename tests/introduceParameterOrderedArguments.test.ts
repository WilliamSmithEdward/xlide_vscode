import { expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { analyzeModule } from '../src/analyzer';
for (const eol of ['\n', '\r\n', '\r']) for (const f of [
 { params:'Optional ByVal mode As Long = 0', next:'Optional ByVal mode As Long = 0, Optional ByVal limit As Long = 3', call:'Report', after:'Report limit:=3' },
 { params:'Optional ByVal mode As Long = 0', next:'Optional ByVal mode As Long = 0, Optional ByVal limit As Long = 3', call:'Report 7', after:'Report 7, limit:=3' },
 { params:'Optional ByVal mode As Long = 0', next:'Optional ByVal mode As Long = 0, Optional ByVal limit As Long = 3', call:'Call Report(7)', after:'Call Report(7, limit:=3)' },
 { params:'Optional ByVal mode As Long = 0', next:'Optional ByVal mode As Long = 0, Optional ByVal limit As Long = 3', call:'Report mode:=7', after:'Report mode:=7, limit:=3' },
 { params:'Optional ByVal first As Long = 0, Optional ByVal second As Long = 1', next:'Optional ByVal first As Long = 0, Optional ByVal second As Long = 1, Optional ByVal limit As Long = 3', call:'Report , 7', after:'Report , 7, limit:=3' },
 { params:'ByVal input As Long, Optional ByVal mode As Long = 0', next:'ByVal input As Long, Optional ByVal mode As Long = 0, Optional ByVal limit As Long = 3', call:'Report 2, 7', after:'Report 2, 7, limit:=3' },
 { params:'ParamArray args() As Variant', next:'ByVal limit As Long, ParamArray args() As Variant', call:'Report', after:'Report 3' },
 { params:'ParamArray args() As Variant', next:'ByVal limit As Long, ParamArray args() As Variant', call:'Report 7, 8', after:'Report 3, 7, 8' },
 { params:'ByVal input As Long, ParamArray args() As Variant', next:'ByVal input As Long, ByVal limit As Long, ParamArray args() As Variant', call:'Call Report(2, 7, 8)', after:'Call Report(2, 3, 7, 8)' },
 { params:'ParamArray args() As Variant', next:'ByVal limit As Long, ParamArray args() As Variant', call:'Report , 7', after:'Report 3,  , 7' },
]) it(`orders ${f.params} with ${f.call} under ${JSON.stringify(eol)}`, () => {
 const source=['Option Explicit',`Public Sub Report(${f.params})`,'Dim limit As Long','limit = 3','Debug.Print limit','End Sub','Sub Caller()',f.call,'End Sub',''].join(eol);
 const result=introduceParameter({source,offset:source.indexOf('limit As'),moduleName:'M'});
 if(!result.ok) throw new Error(result.reason);
 const actual=applyVbaTextEdits(source,result.edits);
 expect(actual).toBe(['Option Explicit',`Public Sub Report(${f.next})`,'Debug.Print limit','End Sub','Sub Caller()',f.after,'End Sub',''].join(eol));
 expect(analyzeModule(actual).filter(d=>['required-param-after-optional','paramarray-not-last','argument-not-optional','named-argument-not-found'].includes(d.code))).toEqual([]);
});
function changed(params: string, call: string, value = '3', type = 'Long', external = false, extra = '') {
 const source=['Option Explicit',`Public Function Report(${params}) As Long`,`Dim limit As ${type}`,`limit = ${value}`,'Report = limit','End Function',extra,external?'':'Sub Caller()',external?'':call,external?'':'End Sub',''].filter(Boolean).join('\n')+'\n';
 const caller='Sub Caller()\n'+call+'\nEnd Sub\n';
 const result=introduceParameter({source,offset:source.indexOf('limit As'),moduleName:'M',otherModuleSources:external?{Caller:caller}:undefined});
 if(!result.ok)return result;
 return {ok:true, source:applyVbaTextEdits(source,result.edits), caller:external?applyVbaTextEdits(caller,result.otherModules![0].edits):''};
}
it('keeps nested calls independently editable',()=>{
 const result=changed('Optional ByVal mode As Long = 0','Debug.Print Report(Report(7))');
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.source).toContain('Debug.Print Report(Report(7, limit:=3), limit:=3)');
});
it('preserves nested commas, strings, parentheses and continued argument text',()=>{
 const result=changed('Optional ByVal mode As Long = 0, Optional ByVal other As Long = 0','Debug.Print Report(Len("a,b"), _\n (2 + 3))');
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.source).toContain('Report(Len("a,b"), _\n (2 + 3), limit:=3)');
});
it('preserves old expression order and appends an effectful initializer last for Optional calls',()=>{
 const result=changed('Optional ByVal mode As Long = 0, Optional ByVal other As Long = 0','Debug.Print Report(NextValue(), other:=NextValue())','NextValue()','Long',false,'Public Function NextValue() As Long\nNextValue = 1\nEnd Function');
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.source).toContain('Report(NextValue(), other:=NextValue(), limit:=NextValue())');
});
it('rewrites cross-module Optional calls using target parameter spelling',()=>{
 const result=changed('Optional ByVal [mode] As Long = 0','Call M.Report(7)','3','Long',true);
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.caller).toContain('Call M.Report(7, limit:=3)');
});
it('preserves omitted Optional slots before the named tail',()=>{
 const result=changed('Optional ByVal a As Long = 0, Optional ByVal b As Long = 0','Debug.Print Report(,)');
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.source).toContain('Report(,, limit:=3)');
});
it('preserves missing ParamArray entries after fixed arguments with spaced commas',()=>{
 const result=changed('ByVal head As Long, ParamArray args() As Variant','Debug.Print Report(2 , , 7)');
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.source).toContain('Report(2 ,3,  , 7)');
});
it('refuses an effectful initializer before existing ParamArray arguments',()=>{
 const result=changed('ParamArray args() As Variant','Debug.Print Report(NextValue())','NextValue()','Long',false,'Public Function NextValue() As Long\nNextValue = 1\nEnd Function');
 expect(result).toMatchObject({ok:false,reason:expect.stringContaining('evaluation order')});
});
it('allows an effectful initializer after fixed arguments when there are no ParamArray entries',()=>{
 const result=changed('ByVal head As Long, ParamArray args() As Variant','Debug.Print Report(2)','NextValue()','Long',false,'Public Function NextValue() As Long\nNextValue = 1\nEnd Function');
 expect(result.ok).toBe(true);
 if(result.ok)expect(result.source).toContain('Report(2, NextValue())');
});
for(const [value,type] of [['-3','Long'],['"a,b"','String'],['True','Boolean'],['#1/1/2000#','Date'],['1.5','Double'],['255','Byte']])it(`allows safe ${type} literal ${value} before ParamArray entries`,()=>expect(changed('ParamArray args() As Variant','Debug.Print Report(7)',value,type).ok).toBe(true));
for(const [value,type] of [['"bad"','Long'],['256','Byte'],['32768','Integer'],['2147483648','Long'],['1 / 0','Long']])it(`refuses unsafe ${type} conversion ${value} before ParamArray entries`,()=>expect(changed('ParamArray args() As Variant','Debug.Print Report(7)',value,type)).toMatchObject({ok:false,reason:expect.stringContaining('conversion')}));
