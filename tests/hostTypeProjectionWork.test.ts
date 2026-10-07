import {expect, it} from 'vitest';
import {typeCompletionCandidates} from '../src/analyzer/completion/typeCompletion';
import {checkMissingLibraryReference} from '../src/analyzer/diagnostics/rules/missingReference';
import type {HostObjectModel, HostType} from '../src/analyzer/host/excelObjectModel';

function measuredModel(hostName = 'Excel') {
    let keyScans = 0, typeReads = 0;
    const types: Record<string, HostType> = {
        'Word.Range': {displayName:'Range', members:[]},
        'Excel.Range': {displayName:'Range', members:[]},
        ...Object.fromEntries(Array.from({length:500},(_,i)=>['Excel.Type'+i,{displayName:'Type'+i,members:[]}])),
    };
    const model: HostObjectModel = {source:'projection work control', hostName, aliases:{}, globals:{},
        enums:{Alignment:{displayName:'Alignment'}},
        types:new Proxy(types, {
            ownKeys(target) { keyScans++; return Reflect.ownKeys(target); },
            get(target,key,receiver) { if(typeof key==='string' && key.includes('.')) { typeReads++; } return Reflect.get(target,key,receiver); },
        }),
    };
    return {model,work:()=>({keyScans,typeReads})};
}

it('reuses host type/enum projections across changing project contexts and caller mutations',()=>{
    const {model,work}=measuredModel();
    const initial=typeCompletionCandidates({model});
    const before=work();
    initial.find(item=>item.name==='Range')!.detail='caller mutation';
    for(let i=0;i<50;i++) {
        const result=typeCompletionCandidates({model,projectTypes:[{name:'Local'+i,kind:'class',moduleName:'Caller'}]});
        expect(result[0].name).toBe('Local'+i);
        expect(result.find(item=>item.name==='Range')!.detail).toBe('Excel type');
        expect(result.some(item=>item.name==='Alignment' && item.kind==='enum')).toBe(true);
    }
    expect(work()).toEqual(before);
    expect(typeCompletionCandidates({model,projectTypes:[{name:'Range',kind:'class',moduleName:'Caller'}]})[0].kind).toBe('class');
    expect(typeCompletionCandidates({model:measuredModel('Word').model}).find(item=>item.name==='Range')!.detail).toBe('Word type');
});

it('reuses library membership for separate module analyses and refreshes with a new model',()=>{
    const {model,work}=measuredModel();
    const check=(input:HostObjectModel, source='Dim app As PowerPoint.Application')=>{
        const diagnostics:string[]=[];
        checkMissingLibraryReference(source,input,rule=>diagnostics.push(rule));
        return diagnostics;
    };
    expect(check(model)).toEqual(['missingLibraryReference']);
    const before=work();
    for(let i=0;i<50;i++) { expect(check(model,'Dim app'+i+' As PowerPoint.Application')).toEqual(['missingLibraryReference']); }
    expect(work()).toEqual(before);
    expect(check({...model,types:{...model.types,'PowerPoint.Application':{displayName:'Application',members:[]}}})).toEqual([]);
});
