import { describe, expect, it } from 'vitest';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { parseModule } from '../src/analyzer/parser/parseModule';

describe('symbol association indexes', () => {
	it('associates procedure symbols with a linear number of AST name-span reads', () => {
		const count = 200;
		const source = 'DefLng A-Z\n' + Array.from({ length: count }, (_, i) =>
			'Public Sub P' + i + '(ByVal arg As Long)\nDim declared As Long\ndeclared = arg\nimplicit = arg\nEnd Sub\n').join('\n');
		const module = parseModule(source);
		let reads = 0;
		const members = module.members.map(member => {
			if (member.kind !== 'Procedure') { return member; }
			const copy = { ...member };
			Object.defineProperty(copy, 'nameSpan', { get() { reads++; return member.nameSpan; } });
			return copy;
		});
		const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: { ...module, members } });
		expect(symbols.implicitLocals?.size).toBe(count);
		for (const locals of symbols.implicitLocals!.values()) { expect([...locals]).toEqual(['implicit']); }
		expect(reads).toBeLessThan(count * 10);
	});

	it('keeps same-name property accessors separate when excluding declared locals', () => {
		const source = 'DefLng A-Z\nPrivate global As Long\n'
			+ 'Public Property Get Item() As Long\nDim getterLocal As Long\ngetterLocal = 1\ngetterImplicit = 1\nglobal = 1\nEnd Property\n'
			+ 'Public Property Let Item(ByVal value As Long)\nDim setterLocal As Long\nsetterLocal = value\nsetterImplicit = value\nEnd Property\n';
		const result = buildModuleSymbols('M', 'class', source);
		expect([...result.implicitLocals!.values()].map(names => [...names])).toEqual([['getterimplicit'], ['setterimplicit']]);
	});

	it('attaches case-insensitive attributes to every matching accessor in source order', () => {
		const source = 'Attribute VB_Name = "M"\nAttribute ITEM.VB_Description = "first"\n'
			+ 'Attribute Item.VB_UserMemId = 0\nAttribute Missing.VB_Description = "ignored"\n'
			+ 'Public Property Get Item() As Long\nEnd Property\n'
			+ 'Public Property Let Item(ByVal value As Long)\nEnd Property\n';
		const result = buildModuleSymbols('M', 'class', source);
		expect(result.root.children).toHaveLength(2);
		for (const symbol of result.root.children!) {
			expect(symbol.attributes?.map(attribute => attribute.name)).toEqual(['VB_Description', 'VB_UserMemId']);
		}
		expect(result.root.attributes?.map(attribute => attribute.name)).toEqual(['VB_Name']);
	});

	it('does not infer locals under Option Explicit or without a DefType directive', () => {
		for (const header of ['Option Explicit\nDefLng A-Z\n', '']) {
			expect(buildModuleSymbols('M', 'standard', header + 'Sub P()\nimplicit = 1\nEnd Sub').implicitLocals).toBeUndefined();
		}
	});
});
