import {describe,expect,it} from 'vitest';
import {assignedValue} from '../src/analyzer/refactor/shared';
import {inlineVariable} from '../src/analyzer/refactor/inlineVariable';
import {introduceParameter} from '../src/analyzer/refactor/introduceParameter';
import {applyVbaTextEdits} from '../src/analyzer/refactor/refactorTypes';
const read=(text:string)=>{const source='prefix\n'+text;return assignedValue(source,{start:7,end:source.length},'limit');};
describe('refactor continued assignment values',()=>{
 for(const eol of ['\n','\r\n','\r']){
  it.each(['limit = _'+eol+' 3','limit _'+eol+' = 3','limit = 3 _'+eol+' + 4','Set limit = _'+eol+' Nothing'])('reads continued assignment %j',text=>{
   expect(read(text)).toBe(text.includes('Nothing')?'Nothing':text.includes('+')?'3  + 4':'3');
  });
  it('inlines a continued literal and removes all assignment lines ('+JSON.stringify(eol)+')',()=>{
   const source=['Sub Go()','Dim limit As Long','limit = _',' 3','Debug.Print limit','End Sub',''].join(eol);
   const result=inlineVariable({source,offset:source.indexOf('limit As')});if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Sub Go()','Debug.Print 3','End Sub',''].join(eol));
  });
  it('introduces a continued literal parameter and updates callers ('+JSON.stringify(eol)+')',()=>{
   const source=['Public Sub Report()','Dim limit As Long','limit = _',' 3','Debug.Print limit','End Sub','Sub Caller()','Report','End Sub',''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit As'),moduleName:'M'});if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal limit As Long)','Debug.Print limit','End Sub','Sub Caller()','Report 3','End Sub',''].join(eol));
  });
  it('does not normalize an uncontinued newline ('+JSON.stringify(eol)+')',()=>expect(read('limit = 3'+eol+' + 4')).toBeUndefined());
 }
 it.each([['limit = 3','3'],['Set limit = Nothing','Nothing'],['limit = "under _ score"','"under _ score"'],['limit = "Rem \' note"','"Rem \' note"'],['limit = 3 \' _ note','3'],['limit = 3: other = 2','3: other = 2'],['other = 3',undefined],['limit =',undefined]])('preserves ordinary value %s',(text,value)=>expect(read(text)).toBe(value));
});
