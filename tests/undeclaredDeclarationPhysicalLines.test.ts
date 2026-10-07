import { expect, it } from 'vitest';
import { checkUndeclaredVariables } from '../src/analyzer/diagnostics/rules/undeclared';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

type Hit = Parameters<Parameters<typeof checkUndeclaredVariables>[13]>;
for(const eol of ['\n','\r\n','\r'])for(const layout of ['plain','declared','continued','block','continuedBlock'])for(const prefixLines of [0,1000]){
 it('keeps declaration inside its procedure '+JSON.stringify({eol,layout,prefixLines}),()=>{
  const prefix=Array(prefixLines).fill("' preceding module line").join(eol)+(prefixLines?eol:'')+'Option Explicit'+eol;
  const header=layout.startsWith('continued')?['Sub Go( _','    ByVal input As Long)'].join(eol):'Sub Go()';
  const nested=layout==='block'||layout==='continuedBlock',declared=layout==='declared';
  const leading=prefix+header+eol+(declared?'    Dim existing As Long'+eol+"    ' retained comment"+eol:'');
  const body=(nested?'    If True Then'+eol:'')+(nested?'        ':'    ')+'missing = 1'+eol+(nested?'    End If'+eol:'');
  const source=leading+body+'End Sub'+eol,module=parseModule(source),symbols=buildModuleSymbols('M','standard',source,{parsedModule:module});
  const hits: Hit[]=[];checkUndeclaredVariables(source,module,symbols,undefined,new Set(),undefined,undefined,undefined,undefined,'standard',undefined,undefined,undefined,(...hit)=>hits.push(hit));
  const fix=hits.find(hit=>hit[3]?.declareVariable)?.[3]?.declareVariable;
  expect(fix).toEqual({variableName:'missing',declaredType:'Long',edit:{span:{start:leading.length,end:leading.length},newText:'    Dim missing As Long'+eol}});
  expect(applyVbaTextEdits(source,[fix!.edit])).toBe(leading+'    Dim missing As Long'+eol+body+'End Sub'+eol);
 });
}
