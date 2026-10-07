import { describe, expect, it } from 'vitest';
import { macroNameTarget } from '../src/analyzer/completion/macroNames';
import type { VbaProcedureSignature } from '../src/analyzer/symbols/symbolModel';
const proc = (name:string,moduleName='Module'):VbaProcedureSignature => ({name,moduleName,kind:'sub',params:[]});
describe('macro string target work',()=>{
  it.each([0,500,999])('stops qualified lookup at its first target, position %i',index=>{
    let reads=0;
    const procedures=Array.from({length:1000},(_,i)=>{const item=proc('P'+i);Object.defineProperty(item,'name',{get(){reads++;return 'P'+i;}});return item;});
    expect(macroNameTarget('Module.P'+index,{macroProcedures:procedures})).toBe(procedures[index]);
    expect(reads).toBeLessThanOrEqual(index+2);
  });
  it('keeps first duplicate, case folding, workbook prefix and external exclusion',()=>{
    const external={...proc('Run','Demo'),external:true},first=proc('Run','Demo'),duplicate=proc('run','DEMO');
    expect(macroNameTarget(" 'Book.xlsm'!dEmO.rUn ",{macroProcedures:[external,first,duplicate]})).toBe(first);
    expect(macroNameTarget('run',{macroProcedures:[external,first,duplicate]})).toBe(first);
  });
  it('requires a unique bare target across modules',()=>{
    const one=proc('Run','A'),two=proc('Run','B');
    expect(macroNameTarget('run',{macroProcedures:[one,two]})).toBeUndefined();
    expect(macroNameTarget('A.Run',{macroProcedures:[one,two]})).toBe(one);
    expect(macroNameTarget('missing',{macroProcedures:[one,two]})).toBeUndefined();
  });
  it('keeps full-name deduplication even for dotted signature names',()=>{
    const odd=proc('Part.Run','Demo'),normal=proc('Run','Demo.Part');
    expect(macroNameTarget('Demo.Part.Run',{macroProcedures:[odd,normal]})).toBeUndefined();
    expect(macroNameTarget('Run',{macroProcedures:[odd,normal]})).toBeUndefined();
    expect(macroNameTarget('Demo.Part.Run',{macroProcedures:[normal,odd]})).toBe(normal);
  });
  it('reads the current caller-owned array and respects an empty macro override',()=>{
    const first=proc('Run','A'),second=proc('Run','B'),list=[first],ctx={projectProcedures:list};
    expect(macroNameTarget('Run',ctx)).toBe(first);
    list.push(second);expect(macroNameTarget('Run',ctx)).toBeUndefined();
    list.splice(0,2,second);expect(macroNameTarget('Run',ctx)).toBe(second);
    expect(macroNameTarget('Run',{...ctx,macroProcedures:[]})).toBeUndefined();
  });
});
