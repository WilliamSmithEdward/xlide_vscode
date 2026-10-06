import {buildModuleSymbols} from '../src/analyzer/symbols/buildModuleSymbols';
import {getExcelObjectModel} from '../src/analyzer/host/excelObjectModel';
import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
import {checkMissingLibraryReference,librariesNamedIn} from '../src/analyzer/diagnostics/rules/missingReference';
const missing=(source:string)=>analyzeProjectModule(source,[],'Caller').filter(d=>d.code==='missing-library-reference');
it.each([
 'Sub T()\nDim Word As Range\nWord.HorizontalAlignment = xlHAlignCenter\nEnd Sub',
 'Sub T(ByVal Word As Range)\nWord.HorizontalAlignment = xlHAlignCenter\nEnd Sub',
 'Private Word As Range\nSub T()\nWord.HorizontalAlignment = xlHAlignCenter\nEnd Sub',
])('does not request Word for a source-bound property write: %s',source=>{
 expect(missing(source)).toEqual([]);
 expect(librariesNamedIn(source).has('word')).toBe(false);
});
it('keeps type and New qualifiers separate from a local value shadow',()=>{
 expect(missing('Sub T()\nDim Word As Range\nDim app As Word.Application\nEnd Sub')).toHaveLength(1);
 expect(missing('Sub T()\nDim Word As Object\nSet Word = New Word.Application\nEnd Sub')).toHaveLength(1);
});
it('does not leak a local shadow into another procedure',()=>{
 const source='Sub T(ByVal Word As Range)\nWord.HorizontalAlignment = xlHAlignCenter\nEnd Sub\nSub U()\nDebug.Print Word.wdFormatPDF\nEnd Sub';
 const errors=missing(source);
 expect(errors).toHaveLength(1);
 expect(source.slice(errors[0].span.start,errors[0].span.end)).toBe('Word.wdFormatPDF');
});
it('recognizes a project-visible value from another module',()=>{
 const source='Sub T()\nWord.HorizontalAlignment = xlHAlignCenter\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Values',source:'Public Word As Range'}],'Caller').filter(d=>d.code==='missing-library-reference')).toEqual([]);
});

it('walks enclosing procedures once across many property writes',()=>{
 const count=500;
 const source=Array.from({length:count},(_,i)=>`Sub T${i}(ByVal Word As Range)\nWord.HorizontalAlignment = xlHAlignCenter\nEnd Sub`).join('\n');
 const symbols=buildModuleSymbols('Caller','standard',source);
 let reads=0;
 symbols.root.children=new Proxy(symbols.root.children!,{get(target,key,receiver){
  if(typeof key==='string' && /^\d+$/.test(key)) { reads++; }
  return Reflect.get(target,key,receiver);
 }});
 const findings:string[]=[];
 checkMissingLibraryReference(source,getExcelObjectModel(),rule=>findings.push(rule),new Set(),{symbols});
 expect(findings).toEqual([]);
 expect(reads).toBeLessThanOrEqual(count*2);
});
