import {expect,it} from 'vitest';
import {projectLibraryDependencies} from '../src/vbaReferenceDependencies';
it('excludes an exported value bound across modules',async()=>{
 expect(await projectLibraryDependencies([{moduleName:'Globals',source:'Public Word As Worksheet'},{moduleName:'Caller',source:'Sub T()\nWord.EnableCalculation = True\nEnd Sub'}],'Word')).toEqual([]);
});
it('excludes a project module qualifier',async()=>{
 expect(await projectLibraryDependencies([{moduleName:'Word',source:'Public value As Long'},{moduleName:'Caller',source:'Sub T()\nWord.value = 20\nEnd Sub'}],'Word')).toEqual([]);
});
it('retains real type and constant dependencies with private foreign shadows',async()=>{
 const modules=[{moduleName:'Globals',source:'Private Word As Worksheet'},{moduleName:'Caller',source:'Sub T()\nDim app As Word.Application\nDebug.Print Word.wdFormatPDF\nEnd Sub'}];
 expect(await projectLibraryDependencies(modules,'Word')).toEqual(['Caller']);
});
it('keeps a type dependency despite a local value shadow',async()=>{
 expect(await projectLibraryDependencies([{moduleName:'Caller',source:'Sub T(ByVal Word As Worksheet)\nDim app As Word.Application\nEnd Sub'}],'Word')).toEqual(['Caller']);
});
