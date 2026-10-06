import {describe,expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
import {createAssignmentCoercionType} from '../src/analyzer/diagnostics/assignmentCoercionType';
function diagnostics(type:string,scope:'variable'|'bare'|'member',value:string) {
 const extra=type === 'Direction'?[{moduleName:'Types',moduleKind:'standard' as const,source:'Public Enum Direction\nNorth = 1\nEnd Enum'}]:[];
 const setter=`Public Property Let State(ByVal value As ${type})\nEnd Property`;
 const body=scope==='variable'?`Dim value As ${type}\nvalue = ${value}\nDebug.Print value`:scope==='bare'?`State = ${value}`:`item.State = ${value}`;
 const source=`Option Explicit\n${scope==='bare'?setter:''}\nSub T(${scope==='member'?'ByVal item As Widget':''})\n${body}\nEnd Sub`;
 return analyzeProjectModule(source,[...extra,...(scope==='member'?[{moduleName:'Widget',moduleKind:'class' as const,source:setter}]:[])],'Caller');
}
describe('enum assignment coercion',()=>{
 for(const type of ['XlHAlign','Excel.XlHAlign','VbMsgBoxResult','VBA.VbMsgBoxResult','Direction']) {
  for(const scope of ['variable','bare','member'] as const) {
   it(`${type} ${scope} rejects nonnumeric text`,()=>expect(diagnostics(type,scope,'"abc"')).toContainEqual(expect.objectContaining({code:'assignment-type-mismatch'})));
   it(`${type} ${scope} accepts an unnamed numeric value`,()=>expect(diagnostics(type,scope,'999').filter(d=>d.severity==='error')).toEqual([]));
   it(`${type} ${scope} checks Long overflow`,()=>expect(diagnostics(type,scope,'3000000000#')).toContainEqual(expect.objectContaining({code:'assignment-type-mismatch',message:expect.stringContaining('Overflow')})));
  }
 }
 it('does not normalize a project class sharing a host enum name',()=>{
  const type=createAssignmentCoercionType({projectClassMembers:[{name:'XlHAlign',moduleName:'XlHAlign',kind:'class',members:[]}]});
  expect(type('XlHAlign')).toBe('XlHAlign');
  expect(type('Excel.XlHAlign')).toBe('Long');
 });
 it('does not normalize a wrong library qualification',()=>expect(createAssignmentCoercionType({})('Missing.XlHAlign')).toBe('Missing.XlHAlign'));
});

it('indexes project types once and does no index work for primitive assignments',()=>{
 let reads=0;
 const type={moduleName:'Types',kind:'enum' as const,members:[],get name(){reads++;return 'Direction';}};
 const normalize=createAssignmentCoercionType({projectClassMembers:[type]});
 for(let i=0;i<1000;i++) expect(normalize('Long')).toBe('Long');
 expect(reads).toBe(0);
 expect(normalize('Direction')).toBe('Long');
 expect(reads).toBeGreaterThan(0);
 reads=0;
 for(let i=0;i<1000;i++) expect(normalize('Direction')).toBe('Long');
 expect(reads).toBe(0);
});

