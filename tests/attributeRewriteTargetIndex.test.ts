import { describe, expect, it, vi } from 'vitest';
import { applyAttributeAnnotations } from '../src/analyzer/annotations/attributeRewriter';
import { PROCEDURE_HEADER, readAttributeAnnotations } from '../src/analyzer/annotations/attributeAnnotations';

describe('attribute rewrite target index', () => {
    it('does not rescan all procedure headers for each annotation', () => {
        const count=100;
        const source='Attribute VB_Name = "Module"\n'+Array.from({length:count},(_,i)=>"'@Description(\"Doc "+i+"\")\n'@ExcelHotkey(\"A\")\nSub P"+i+'()\n'+Array.from({length:10},()=> 'Debug.Print 1').join('\n')+'\nEnd Sub\n').join('');
        const annotations=readAttributeAnnotations(source);
        let reads=0;
        const original=PROCEDURE_HEADER.exec;
        const spy=vi.spyOn(PROCEDURE_HEADER,'exec').mockImplementation(function(this: RegExp, text: string) {
            reads++; return original.call(this,text);
        });
        try {
            const result=applyAttributeAnnotations(source,annotations);
            expect(result.changes).toHaveLength(count*2);
            expect(result.skipped).toEqual([]);
            expect(reads).toBeLessThan(source.split('\n').length*5);
        } finally { spy.mockRestore(); }
    });

    it.each(['\n','\r\n'])('keeps repeated edits, variable scope, and insertion order with %j', eol => {
        const source=['Attribute VB_Name = "Module"','Private x As Long', 'Sub P()', 'Attribute Other.VB_Description = "leave"', 'End Sub', 'Private late As Long', ''].join(eol);
        const result=applyAttributeAnnotations(source,{problems:[],annotations:[
            {kind:'Description',line:1,target:'p',argument:'one'},
            {kind:'ExcelHotkey',line:1,target:'P',argument:'A'},
            {kind:'Description',line:1,target:'P',argument:'two'},
            {kind:'ModuleDescription',line:1,argument:'module'},
            {kind:'VariableDescription',line:1,target:'x',argument:'variable'},
            {kind:'VariableDescription',line:1,target:'late',argument:'skip'},
        ]});
        expect(result.changes.map(change=>[change.target,change.attribute,change.from,change.to])).toEqual([
            ['p','VB_Description',undefined,'"one"'],
            ['P','VB_ProcData.VB_Invoke_Func',undefined,'"A\\n14"'],
            ['P','VB_Description','"one"','"two"'],
            ['module','VB_Description',undefined,'"module"'],
            ['x','VB_VarDescription',undefined,'"variable"'],
        ]);
        expect(result.text).toContain(['Sub P()', 'Attribute P.VB_ProcData.VB_Invoke_Func = "A\\n14"','Attribute P.VB_Description = "two"','Attribute Other.VB_Description = "leave"'].join(eol));
        expect(result.skipped).toEqual(["no module-level variable named 'late' was found for '@VariableDescription."]);
    });
});
