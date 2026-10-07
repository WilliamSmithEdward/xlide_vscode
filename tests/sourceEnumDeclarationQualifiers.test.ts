import {expect,it} from 'vitest';
import {analyzeModule} from '../src/analyzer';
import {analyzeProjectModule} from './diagnostics/helpers';
import {resolveAssignmentValueCompletion} from '../src/analyzer/completion/assignmentValueCompletion';
it('rejects a source module qualifier in an enum As declaration',()=>{
 const source='Sub T()\nDim value As Types.Direction\nvalue = Types.Direction.North\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Types',source:'Public Enum Direction\nNorth = 1\nEnd Enum'}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'invalid-as-type-name'})]);
});
it('keeps an unqualified enum declaration and qualified constant valid',()=>{
 const source='Sub T()\nDim value As Direction\nvalue = Types.Direction.North\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Types',source:'Public Enum Direction\nNorth = 1\nEnd Enum'}],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
it('does not suggest assignment values for an invalid own-module enum type',()=>{
 const source='Enum Direction\nNorth = 1\nEnd Enum\nSub T()\nDim value As Module.Direction\nvalue = ';
 expect(resolveAssignmentValueCompletion(source,source.length)).toBeUndefined();
});

it.each([['Excel','XlHAlign','xlHAlignCenter'],['VBA','VbMsgBoxResult','vbYes']])('keeps real library qualification separate from a module named %s',(moduleName,enumName,constant)=>{
 const source='Enum '+enumName+'\nBogus = 1\nEnd Enum\nSub T()\nDim value As '+moduleName+'.'+enumName+'\nvalue = ';
 expect(resolveAssignmentValueCompletion(source,source.length,{moduleName})?.constants.map(c=>c.name)).toContain(constant);
});

it('checks a standalone own-module qualifier without project metadata',()=>{
 const source='Enum Direction\nNorth = 1\nEnd Enum\nSub T()\nDim value As Module.Direction\nEnd Sub';
 expect(analyzeModule(source,{moduleName:'Module'}).filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'invalid-as-type-name'})]);
});
it('preserves a qualified source UDT declaration',()=>{
 const source='Sub T()\nDim value As Types.Record\nvalue.Flag = True\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Types',source:'Public Type Record\nFlag As Boolean\nEnd Type'}],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
