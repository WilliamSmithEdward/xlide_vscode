import {expect,it} from 'vitest';
import {tokenizeCached} from '../src/analyzer/lexer/tokenize';
import {inferArgumentType,unwrapOuterParens} from '../src/analyzer/diagnostics/typeInference';
it('infers deeply grouped assignment values without overflowing the stack',()=>{
 const source='('.repeat(4000)+'True'+')'.repeat(4000);
 expect(inferArgumentType(tokenizeCached(source),0,new Map(),new Map())?.type).toBe('Boolean');
});
it('unwraps all complete outer groups while preserving inner operand groups',()=>{
 const tokens=tokenizeCached('(((1)+(2)))');
 expect(unwrapOuterParens(tokens).map(t=>t.rawText).join('')).toBe('(1)+(2)');
});
it('reads grouped tokens a linear number of times',()=>{
 let reads=0;
 const tokens=tokenizeCached('('.repeat(10000)+'True'+')'.repeat(10000)).map(t=>new Proxy(t,{get(target,key,receiver){if(key==='rawText')reads++;return Reflect.get(target,key,receiver);}}));
 expect(unwrapOuterParens(tokens).map(t=>t.rawText)).toEqual(['True']);
 expect(reads).toBeLessThan(tokens.length*2);
});
it.each(['(1)+(2)','((1)','(1))','1',''])('preserves incomplete or unenclosed expressions: %s',source=>{
 const tokens=tokenizeCached(source);
 expect(unwrapOuterParens(tokens)).toBe(tokens);
});
