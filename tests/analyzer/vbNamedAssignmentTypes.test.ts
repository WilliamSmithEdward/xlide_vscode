import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
it('does not collapse VbWidget array identity into Widget',()=>{
 const source='Function Factory() As VbWidget()\nDim data(1) As VbWidget\nFactory = data\nEnd Function\nSub T()\nDim values() As Widget\nvalues = Factory()\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:'Option Explicit'},{moduleName:'VbWidget',moduleKind:'class',source:'Option Explicit'}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'array-target-assignment'})]);
});
it('does not give a VbCollection class the runtime Collection default',()=>{
 const source='Sub T()\nDim item As VbCollection\nSet item = New VbCollection\nitem = 20\nEnd Sub';
 const errors=analyzeProjectModule(source,[{moduleName:'VbCollection',moduleKind:'class',source:'Option Explicit'}],'Caller').filter(d=>d.severity==='error');
 expect(errors).toHaveLength(1);
 expect(errors[0]?.message).toContain("Run-time error '438'");
});
it('keeps VbInteger source enum storage as Long',()=>{
 const source='Enum VbInteger\nFirst = 1\nEnd Enum\nProperty Let State(ByVal value As VbInteger)\nEnd Property\nSub T()\nState = 50000\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
