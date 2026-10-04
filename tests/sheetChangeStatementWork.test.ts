import { afterEach, expect, it, vi } from 'vitest';
const work=vi.hoisted(()=>({reads:0,tokens:0}));
vi.mock('../src/analyzer/lexer/tokenize',async original=>{const actual=await original<typeof import('../src/analyzer/lexer/tokenize')>();return{...actual,tokenizeCached:(source:string)=>{const tokens=actual.tokenizeCached(source);work.tokens=tokens.length;return tokens.map(token=>({...token,get rawText(){work.reads++;return token.rawText;}}));}};});
import { sheetChangesIn } from '../src/analyzer/symbols/sheetChanges';
afterEach(()=>vi.restoreAllMocks());
for(const count of [1,100,1000])it('classifies range words once for '+count+' Copy members',()=>{
 const source='Take '+Array.from({length:count},(_,i)=>'o'+i+'.Copy(1)').join(', ');work.reads=0;
 expect(sheetChangesIn(source)).toEqual({addsSheets:true,namesAssigned:new Set(),assignsComputedName:false});
 expect(work.reads).toBeLessThanOrEqual(30*work.tokens);
});
for(const count of [1,100,1000])it('does not copy source tails for '+count+' Name comparisons',()=>{
 const source='If '+Array.from({length:count},(_,i)=>'o'+i+'.Name = value'+i).join(' And ')+' Then Debug.Print 1';let copied=0;const original=Array.prototype.slice;
 vi.spyOn(Array.prototype,'slice').mockImplementation(function(this:unknown[],start?:number,end?:number){if((this[0] as {rawText?:string})?.rawText==='If'){const at=start??0,limit=end??this.length;copied+=Math.max(0,limit-at);}return original.call(this,start,end);});
 expect(sheetChangesIn(source)).toEqual({addsSheets:false,namesAssigned:new Set(),assignsComputedName:false});
 expect(copied).toBeLessThanOrEqual(10*work.tokens);
});
