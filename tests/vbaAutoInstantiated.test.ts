import { describe, expect, it } from 'vitest';
import { ProjectIndex } from '../src/analyzer';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

// Issue #16. `Dim x As New Invoice` is instantiated by VBA on ANY access -
// including the first, and including after `Set x = Nothing` - so member access
// on it can never raise run-time error 91. The analyzer reported one anyway,
// because the symbol did not record that the declaration said New.
function projectWithInvoice(): ProjectIndex {
	const index = new ProjectIndex();
	index.setModule({
		moduleName: 'Invoice',
		moduleKind: 'class',
		source: [
			'Option Explicit',
			'Private mNumber As String',
			'',
			'Public Property Let Number(ByVal value As String)',
			'    mNumber = value',
			'End Property',
		].join('\n'),
	});
	index.setModule({ moduleName: 'Builder', moduleKind: 'standard', source: '' });
	return index;
}

function codesFor(lines: string[]): string[] {
	const index = projectWithInvoice();
	return analyzeVbaModuleSource({
		source: lines.join('\r\n'),
		moduleName: 'Builder',
		moduleType: 'standard',
		moduleKind: 'standard',
		projectClassMembers: index.projectMemberSurfaces('Builder'),
	} as never).diagnostics.map((d) => String(d.code));
}

describe('a variable declared As New', () => {
	it('is not reported as Nothing before member access', () => {
		expect(codesFor([
			'Option Explicit',
			'',
			'Public Function MakeInvoice(ByVal number As String) As Invoice',
			'    Dim made As New Invoice',
			'    made.Number = number',
			'    Set MakeInvoice = made',
			'End Function',
			'',
		])).not.toContain('object-variable-not-set');
	});

	it('is still not reported after being set to Nothing', () => {
		// The VBA gotcha: touching a member re-instantiates it, so this runs.
		expect(codesFor([
			'Option Explicit',
			'',
			'Public Sub Revive()',
			'    Dim revived As New Invoice',
			'    Set revived = Nothing',
			'    revived.Number = "back again"',
			'End Sub',
			'',
		])).not.toContain('object-variable-not-set');
	});
});

describe('the finding this rule exists for still fires', () => {
	it('reports a plain Dim whose member is touched before any Set', () => {
		expect(codesFor([
			'Option Explicit',
			'',
			'Public Sub NeverSet()',
			'    Dim made As Invoice',
			'    made.Number = "raises 91"',
			'End Sub',
			'',
		])).toContain('object-variable-not-set');
	});

	it('stays quiet when the variable was Set first', () => {
		expect(codesFor([
			'Option Explicit',
			'',
			'Public Sub SetBeforeUse()',
			'    Dim made As Invoice',
			'    Set made = New Invoice',
			'    made.Number = "fine"',
			'End Sub',
			'',
		])).not.toContain('object-variable-not-set');
	});
});

describe('the symbol records how the variable was declared', () => {
	it('distinguishes As New from a plain declaration', () => {
		// The two were indistinguishable before: both carried only asType.
		const source = 'Public Sub T()\r\n    Dim made As New Invoice\r\n    Dim plain As Invoice\r\nEnd Sub\r\n';
		const symbols = buildModuleSymbols(parseModule(source), 'M', source);
		const locals: Array<{ name: string; isAutoInstantiated?: boolean }> = [];
		const walk = (sym: { kind: string; name: string; isAutoInstantiated?: boolean; children?: unknown[] }): void => {
			if (sym.kind === 'localVariable') {
				locals.push({ name: sym.name, isAutoInstantiated: sym.isAutoInstantiated });
			}
			for (const child of (sym.children ?? []) as never[]) { walk(child); }
		};
		walk(symbols.root as never);
		expect(locals.find((l) => l.name === 'made')?.isAutoInstantiated).toBe(true);
		expect(locals.find((l) => l.name === 'plain')?.isAutoInstantiated).toBe(false);
	});
});
