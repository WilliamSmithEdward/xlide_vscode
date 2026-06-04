import { describe, expect, it } from 'vitest';
import { ProjectIndex } from '../src/analyzer';
import {
	documentOutlineSymbols,
	documentOutlineSymbolsForSource,
	workspaceSymbols,
} from '../src/vbaSymbolPresentation';

describe('VBA symbol presentation', () => {
	it('builds hierarchical document symbols from the AST project graph', () => {
		const source = [
			'Private mState As Long',
			'Public Const MaxItems As Long = 10',
			'',
			'Public Type TPoint',
			'    X As Long',
			'    Y As Long',
			'End Type',
			'',
			'Public Sub DoWork(ByVal count As Long)',
			'    Dim total As Long',
			'End Sub',
		].join('\n');

		const symbols = documentOutlineSymbolsForSource('Module1', 'standard', source);
		expect(symbols.map((symbol) => `${symbol.name}:${symbol.kind}:${symbol.detail}`)).toEqual([
			'mState:moduleVariable:Variable As Long',
			'MaxItems:constant:Const As Long',
			'TPoint:type:Type',
			'DoWork:sub:Sub',
		]);

		expect(symbols.find((symbol) => symbol.name === 'TPoint')?.children
			.map((child) => `${child.name}:${child.kind}:${child.detail}`))
			.toEqual(['X:typeField:Field As Long', 'Y:typeField:Field As Long']);
		expect(symbols.find((symbol) => symbol.name === 'DoWork')?.children
			.map((child) => `${child.name}:${child.kind}:${child.detail}`))
			.toEqual(['count:parameter:Parameter As Long', 'total:localVariable:Local As Long']);
	});

	it('uses the project workspace symbol filter without module roots', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Invoices',
			moduleKind: 'standard',
			source: 'Public Sub InvoiceTotal()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Customers',
			moduleKind: 'class',
			source: 'Public Function InvoiceCaption() As String\nEnd Function\n',
		});

		const symbols = workspaceSymbols(index, 'invoice')
			.map((symbol) => `${symbol.name}:${symbol.kind}:${symbol.containerName}`);

		expect(symbols).toEqual([
			'InvoiceTotal:sub:Invoices',
			'InvoiceCaption:function:Customers',
		]);
		expect(workspaceSymbols(index, 'module')).toEqual([]);
	});

	it('keeps enum members and UDT fields nested for document outlines', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: [
				'Public Enum Status',
				'    Draft',
				'    Posted',
				'End Enum',
				'Public Type InvoiceLine',
				'    Amount As Currency',
				'End Type',
			].join('\n'),
		});

		const symbols = documentOutlineSymbols(index, 'Types');
		expect(symbols.find((symbol) => symbol.name === 'Status')?.children
			.map((child) => `${child.name}:${child.kind}:${child.containerName}`))
			.toEqual(['Draft:enumMember:Status', 'Posted:enumMember:Status']);
		expect(symbols.find((symbol) => symbol.name === 'InvoiceLine')?.children
			.map((child) => `${child.name}:${child.kind}:${child.containerName}`))
			.toEqual(['Amount:typeField:InvoiceLine']);
	});
});
