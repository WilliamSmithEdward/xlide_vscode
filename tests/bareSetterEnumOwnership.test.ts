import {afterEach,describe,expect,it,vi} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
import {resolveAssignmentValueCompletion} from '../src/analyzer/completion/assignmentValueCompletion';
import {getExcelObjectModel} from '../src/analyzer/host/excelObjectModel';
import {getWordObjectModel} from '../src/analyzer/host/wordObjectModel';
import {getPowerPointObjectModel} from '../src/analyzer/host/powerpointObjectModel';
import {getAccessObjectModel} from '../src/analyzer/host/accessObjectModel';
import {hostObjectModelForTokens} from '../src/analyzer/host/hostRegistry';
import {resolveHostEnum} from '../src/analyzer/host/hostModel';
function diagnostics(statement:string,indexed=false,type='Boolean',extra='') {
 const setter=`Public Property Let State(${indexed?'ByVal index As Long, ':''}ByVal value As ${type})\nEnd Property`;
 return analyzeProjectModule(`Option Explicit\n${setter}\nSub T()\n${extra}\n${statement}\nEnd Sub`,[],'Caller');
}
describe('bare Property Let value compatibility',()=>{
 it.each(['State = "nonsense"','Let State = "nonsense"','If True Then State = "nonsense"'])('flags invalid Boolean values: %s',statement=>{
  expect(diagnostics(statement)).toEqual([expect.objectContaining({code:'assignment-type-mismatch',severity:'error'})]);
 });
 it('checks indexed bare setters',()=>expect(diagnostics('State(1) = "nonsense"',true)).toEqual([expect.objectContaining({code:'assignment-type-mismatch'})]));
 it.each(['True','1','"True"'])('allows Boolean coercion %s',value=>expect(diagnostics('State = '+value)).toEqual([]));
 it('uses the local variable shadow instead of the setter',()=>expect(diagnostics('State = "text"\nDebug.Print State',false,'Boolean','Dim State As String')).toEqual([]));
 it('rejects a known nonnumeric String for a numeric setter',()=>expect(diagnostics('State = "nonsense"',false,'Long')).toContainEqual(expect.objectContaining({code:'assignment-type-mismatch'})));
 it('ignores comparisons and reads',()=>expect(diagnostics('If State = True Then Exit Sub',false,'Boolean','Dim State As Boolean')).toEqual([]));
});
describe('shared enum ownership',()=>{
 it.each([getExcelObjectModel,getWordObjectModel,getPowerPointObjectModel,getAccessObjectModel])('keeps Office ownership in a single host',factory=>expect(resolveHostEnum('MsoTriState',factory())?.library).toBe('Office'));
 it('keeps shared ownership when host references are merged',()=>expect(resolveHostEnum('MsoTriState',hostObjectModelForTokens(['excel','word']))?.library).toBe('Office'));
 it('keeps DAO ownership inside Access',()=>expect(resolveHostEnum('RecordsetTypeEnum',getAccessObjectModel())?.library).toBe('DAO'));
 it('offers the correct qualification for a shadowed Office constant',()=>{
  const source='Sub T(ByVal sh As Shape)\nsh.Visible = ';
  expect(resolveAssignmentValueCompletion(source,source.length)?.qualifiedEnumName).toBe('Office.MsoTriState');
 });
});


afterEach(()=>vi.restoreAllMocks());
it('does not parse ordinary assignment targets merely because the module has a setter',async()=>{
 const syntax=await import('../src/analyzer/completion/assignmentTarget');
 const parse=vi.spyOn(syntax,'assignmentTargetFromTokens');
 const source='Option Explicit\nProperty Let State(ByVal value As Boolean)\nEnd Property\nSub WorkProbe()\nDim x As Long\n'+'x = 1\n'.repeat(1000)+'State = True\nDebug.Print x\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
 expect(parse).toHaveBeenCalledTimes(1);
});


it('uses only the active conditional setter declaration',()=>{
 const source='Option Explicit\n#If False Then\nProperty Let State(ByVal value As Boolean)\nEnd Property\n#Else\nProperty Let State(ByVal value As String)\nEnd Property\n#End If\nSub T()\nState = "nonsense"\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});

it('does not flag the valid qualified Office constant as undeclared',()=>{
 const source='Option Explicit\nSub T(ByVal sh As Shape)\nDim msoTrue As Long\nsh.Visible = Office.MsoTriState.msoTrue\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});

it('reuses library namespace metadata across module analyses',()=>{
 const base=getExcelObjectModel(); let scans=0;
 const types=new Proxy(base.types,{ownKeys(target){if(new Error().stack?.includes('libraryQualifierNames'))scans++;return Reflect.ownKeys(target);}});
 const model={...base,types};
 const source='Option Explicit\nSub T()\nDim value As Long\nvalue = Office.MsoTriState.msoTrue\nDebug.Print value\nEnd Sub';
 analyzeProjectModule(source,[],'Warm',{hostModel:model});
 expect(scans).toBeGreaterThan(0);
 scans=0;
 for(let i=0;i<10;i++) analyzeProjectModule(source.replace('Sub T()',`Sub T${i}()`),[],'Module'+i,{hostModel:model});
 expect(scans).toBe(0);
});
