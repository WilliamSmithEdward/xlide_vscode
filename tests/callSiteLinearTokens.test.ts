import { describe, expect, it, vi } from 'vitest';
const work=vi.hoisted(()=>({characters:0,calls:0}));
vi.mock('../src/analyzer/lexer/tokenize',async importOriginal=>{
 const actual=await importOriginal<typeof import('../src/analyzer/lexer/tokenize')>();
 const measured=(source:string)=>{
  work.characters+=source.length;work.calls++;
  const tokens=actual.tokenize(source);
  for(const token of tokens){for(const trivia of token.leadingTrivia??[])Object.freeze(trivia);if(token.leadingTrivia)Object.freeze(token.leadingTrivia);Object.freeze(token);}
  return Object.freeze(tokens);
 };
 return {...actual,tokenize:measured};
});
import { tokenizeCached, startTokenizeMissLogForTests, stopTokenizeMissLogForTests } from '../src/analyzer/lexer/tokenize';
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
describe('call-site lexer work is bounded with immutable cached tokens',()=>{
 it('does not evict module token caches with refactor fragments',()=>{
  const source='Sub Caller()\n'+Array.from({length:1000},(_,i)=>'Go "value '+i+'": Debug.Print '+i).join('\n')+'\nEnd Sub';
  const moduleTokens=tokenizeCached(source);
  startTokenizeMissLogForTests();
  const sites=callSitesOf(source,'Go');
  expect(sites).toHaveLength(1000);
  expect(tokenizeCached(source)).toBe(moduleTokens);
  expect(stopTokenizeMissLogForTests()).toEqual([]);
 });

 for(const count of [1,100,1000]){
  it('tokenizes a colon-separated line once at '+count+' calls',()=>{
   const body=Array(count).fill('Go 1').join(': ');
   const source='Sub Caller()\n'+body+'\nEnd Sub';
   work.characters=0;work.calls=0;
   const sites=callSitesOf(source,'Go');
   expect(sites).toHaveLength(count);
   expect(work.characters).toBeLessThanOrEqual(source.length);
   expect(work.calls).toBeLessThanOrEqual(1);
   const output=applyVbaTextEdits(source,sites.map(s=>({span:s.argumentInsert,newText:s.argumentText('3')})));
   expect(output).toBe(source.replace(body,Array(count).fill('Go 1, 3').join(': ')));
  });
  it('tokenizes a nested call line once at '+count+' calls',()=>{
   const body='x = '+'Go('.repeat(count)+'1'+')'.repeat(count);
   const source='Sub Caller()\n'+body+'\nEnd Sub';
   work.characters=0;work.calls=0;
   const sites=callSitesOf(source,'Go');
   expect(sites).toHaveLength(count);
   expect(work.characters).toBeLessThanOrEqual(source.length);
   expect(work.calls).toBeLessThanOrEqual(1);
   const output=applyVbaTextEdits(source,sites.map(s=>({span:s.argumentInsert,newText:s.argumentText('3')})));
   expect(output).toBe(source.replace(body,'x = '+'Go('.repeat(count)+'1'+', 3)'.repeat(count)));
  });
 }
});
