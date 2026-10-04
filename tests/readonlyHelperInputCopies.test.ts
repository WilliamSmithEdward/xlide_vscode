import { expect, it } from 'vitest';
import { checkRuntimeMemberNotFound } from '../src/analyzer/diagnostics/rules/lateBoundMembers';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { statementTokensCached } from '../src/analyzer/lexer/tokenHelpers';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';

const message = 'CreateObject: "Dictionary" lacks its library: the ProgID is "scripting.dictionary". This will raise Run-time error \'429\': ActiveX component can\'t create object.';
function fixture(extra = '', progId = 'Dictionary') {
	const source = ['Option Explicit', 'Sub Go()', 'Dim obj As Object', 'Set obj = CreateObject("' + progId + '"' + extra + ')', 'End Sub', ''].join('\n');
	const mod = parseModule(source), procedure = mod.members.find(m => m.kind === 'Procedure')!;
	const statement = procedure.body.find(node => source.slice(node.span.start, node.span.end).includes('CreateObject'))!;
	const tokens = statementTokensCached(source, statement.span), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	return { source, mod, tokens, symbols };
}
function expected(source: string, code = 'runtimeArgumentValue') {
	const start = source.indexOf('"Dictionary"'); return [[code, message, { start, end: start + 12 }]];
}
function run(f: ReturnType<typeof fixture>) {
	const out: unknown[][] = [];
	checkRuntimeMemberNotFound(f.source, f.mod, f.symbols, { parsedModule: f.mod, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v));
	return out;
}
it.each([10, 100, 1000])('avoids whole-array helper copies with %i extra call arguments', count => {
	const f = fixture(', ' + Array(count).fill('0').join(', ')), original = Object.getOwnPropertyDescriptor(f.tokens, Symbol.iterator);
	let yields = 0;
	Object.defineProperty(f.tokens, Symbol.iterator, { configurable: true, value: function* (this: typeof f.tokens) { for (let i = 0; i < this.length; i++) { yields++; yield this[i]; } } });
	try { expect(run(f)).toEqual(expected(f.source)); expect(yields).toBeLessThanOrEqual(f.tokens.length); }
	finally { if (original) Object.defineProperty(f.tokens, Symbol.iterator, original); else delete (f.tokens as unknown as Record<symbol, unknown>)[Symbol.iterator]; }
});
it.each(['Dictionary', 'Scripting.Dictionary'])('accepts frozen shared tokens for public and full diagnostics: %s', progId => {
	const f = fixture('', progId), snapshot = f.tokens.map(t => ({ ...t }));
	for (const token of f.tokens) Object.freeze(token); Object.freeze(f.tokens);
	expect(run(f)).toEqual(progId === 'Dictionary' ? expected(f.source) : []);
	const errors: unknown[] = [];
	expect(analyzeModule(f.source, { onInternalError: e => errors.push(e) }).map(d => [d.code, d.message, d.span])).toEqual([...(progId === 'Dictionary' ? expected(f.source, 'runtime-argument-value') : []), ['variable-never-read', "Variable 'obj' is assigned but its value is never read.", { start: f.source.indexOf('Dim obj') + 4, end: f.source.indexOf('Dim obj') + 7 }]]);
	expect(errors).toEqual([]); expect(f.tokens).toEqual(snapshot);
});
