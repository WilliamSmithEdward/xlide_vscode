import { describe, expect, it } from 'vitest';
import { projectTypeAt, projectClassMemberAt, privateMemberOwnerAt, type MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
const source='Option Explicit\nSub Main()\nDim c As C\nDebug.Print c.Value\nEnd Sub\n';
const offset=source.indexOf('c.Value')+2;
const surface=(name='C',kind:VbaProjectClassMembers['kind']='class'):VbaProjectClassMembers=>({name,moduleName:name,kind,members:[{name:'Value',kind:'property',returns:'Long',moduleName:name}],privateMembers:['Hidden']});
const lookups=[
    {name:'project type',lookup:(ctx:MemberCompletionContext)=>projectTypeAt(source,offset,ctx)},
    {name:'class member',lookup:(ctx:MemberCompletionContext)=>projectClassMemberAt(source,offset,'Value',ctx)},
    {name:'private owner',lookup:(ctx:MemberCompletionContext)=>privateMemberOwnerAt(source,offset,'Hidden',ctx)},
];
describe.each(lookups)('$name with no indexed project surfaces',({lookup})=>{
    it.each(['absent','empty','ambiguous'])('avoids receiver token-prefix work for %s metadata',mode=>{
        let reads=0;
        const ctx:MemberCompletionContext={
            ...(mode==='absent'?{}:{projectClassMembers:mode==='empty'?[]:[surface('C'),surface('c')]}),
            get sourceTokens(){reads++;return undefined;},
        };
        for(let i=0;i<1000;i++)expect(lookup(ctx)).toBeUndefined();
        expect(reads).toBe(0);
    });
});
it('resolves indexed class members, case-insensitive first matches and private owners',()=>{
    const type=surface();type.members.push({...type.members[0],name:'VALUE',returns:'String'});
    const ctx={projectClassMembers:[type]};
    expect(projectTypeAt(source,offset,ctx)).toBe(type);
    expect(projectClassMemberAt(source,offset,'vAlUe',ctx)).toBe(type.members[0]);
    expect(privateMemberOwnerAt(source,offset,'hidden',ctx)).toBe('C');
    expect(privateMemberOwnerAt(source,offset,'Value',ctx)).toBeUndefined();
});
it('uses a new metadata list after an empty lookup with the same context',()=>{
    const ctx:MemberCompletionContext={projectClassMembers:[]};
    expect(projectTypeAt(source,offset,ctx)).toBeUndefined();
    const type=surface();ctx.projectClassMembers=[type];
    expect(projectTypeAt(source,offset,ctx)).toBe(type);
    expect(projectClassMemberAt(source,offset,'Value',ctx)).toBe(type.members[0]);
    expect(privateMemberOwnerAt(source,offset,'Hidden',ctx)).toBe('C');
    ctx.projectClassMembers=[];
    expect(projectTypeAt(source,offset,ctx)).toBeUndefined();
});
it('still resolves non-class surfaces without making them class members',()=>{
    const type=surface('C','standardModule'),ctx={projectClassMembers:[type]};
    const text='Option Explicit\nSub Main()\nDebug.Print C.Value\nEnd Sub\n',at=text.indexOf('C.Value')+2;
    expect(projectTypeAt(text,at,ctx)).toBe(type);
    expect(projectClassMemberAt(text,at,'Value',ctx)).toBeUndefined();
});
