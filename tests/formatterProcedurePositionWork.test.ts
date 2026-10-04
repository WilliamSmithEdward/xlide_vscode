import {afterEach,describe,expect,it,vi} from 'vitest';
import {formatVbaModule} from '../src/analyzer/format/formatModule';
afterEach(()=>vi.restoreAllMocks());
describe('formatter procedure position work',()=>{
 for(const eol of ['\n','\r\n','\r'])it.each([10,100,500])('avoids depth-sized procedure scans (%i, '+JSON.stringify(eol)+')',depth=>{
  const source='#If FLAG Then'+eol;const text=source.repeat(depth)+'Debug.Print 1'+eol+('#End If'+eol).repeat(depth);
  const lines=[...Array.from({length:depth},(_,i)=>' '.repeat(4*i)+'#If FLAG Then'),' '.repeat(4*depth)+'Debug.Print 1',...Array.from({length:depth},(_,i)=>' '.repeat(4*(depth-i-1))+'#End If')];
  let visits=0;const original=Array.prototype.some;
  vi.spyOn(Array.prototype,'some').mockImplementation(function(this:unknown[],callback,thisArg){return original.call(this,(value,index,array)=>{const block=value as {kind?:string;level?:number;moduleLevel?:boolean};if(block&&typeof block.level==='number'&&typeof block.moduleLevel==='boolean')visits++;return callback.call(thisArg,value,index,array);});});
  expect(formatVbaModule(text,{tabSize:4,insertSpaces:true})).toEqual({text:lines.join(eol)+eol});
  expect(visits).toBeLessThanOrEqual(depth*2);
 });
 it.each(['\n','\r\n','\r'])('resets procedure ownership after module branch truncation (%j)',eol=>{
  const source=['#If FLAG Then','Sub First()','If FLAG Then','#Else','Sub Second()','Debug.Print 1','End Sub','#End If'].join(eol);
  const expected=['#If FLAG Then','    Sub First()','        If FLAG Then','#Else','    Sub Second()','        Debug.Print 1','    End Sub','#End If'].join(eol);
  expect(formatVbaModule(source,{tabSize:4,insertSpaces:true})).toEqual({text:expected});
 });
 it.each(['\n','\r\n','\r'])('resets after a mismatched closer removes the procedure (%j)',eol=>{
  const source=['#If FLAG Then','Sub First()','#End If','#If OTHER Then','Sub Second()','Debug.Print 1','End Sub','#End If'].join(eol);
  const expected=['#If FLAG Then','    Sub First()','#End If','#If OTHER Then','    Sub Second()','        Debug.Print 1','    End Sub','#End If'].join(eol);
  expect(formatVbaModule(source,{tabSize:4,insertSpaces:true})).toEqual({text:expected});
 });
 it('keeps only module-level conditional blocks at replacement and stray procedure boundaries',()=>{
  const source='#If FLAG Then\nSub First()\n#If INNER Then\nSub Second()\nDebug.Print 1\nEnd Sub\n#End If\nEnd Sub\n#If OTHER Then\nSub Last()\nEnd Sub\n#End If';
  expect(formatVbaModule(source,{tabSize:4,insertSpaces:true})).toEqual({text:'#If FLAG Then\n    Sub First()\n        #If INNER Then\n    Sub Second()\n        Debug.Print 1\n    End Sub\n#End If\nEnd Sub\n#If OTHER Then\n    Sub Last()\n    End Sub\n#End If'});
 });
});
