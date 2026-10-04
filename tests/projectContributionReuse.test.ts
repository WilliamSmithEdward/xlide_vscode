import { expect, it } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';

const source = (name: string) => `Public Const ${name}Value As Long = 7\nPublic Type ${name}Type\nField As Long\nEnd Type\nPublic Enum ${name}Enum\n${name}Item\nEnd Enum\nPublic Sub ${name}Public()\nEnd Sub\nPrivate Sub ${name}Private()\nEnd Sub\n`;
function queries(index: ProjectIndex, name: string) {
	return {
		procedures: index.visibleProcedureNames(name),
		signatures: index.visibleProcedureSignatures(name),
		identifiers: index.visibleIdentifierNames(name),
		symbols: index.visibleIdentifierSymbols(name),
		nonTypes: index.visibleNonTypeNames(name),
		types: index.visibleTypeNames(name),
		constants: index.visibleExternalIntegerConstantExpressions(name),
		standard: index.projectStandardModuleMembers(name),
		surfaces: index.projectMemberSurfaces(name),
	};
}
for (const count of [2, 20, 100]) {
	it(`reuses unchanged contributions across three edits in a ${count}-module project`, () => {
		const index = new ProjectIndex();
		for (let i = 0; i < count; i++) index.setModule({ moduleName: `M${i}`, moduleKind: 'standard', source: source(`N${i}`) });
		let reads = 0;
		for (let i = 1; i < count; i++) {
			const root = index.getModule(`M${i}`)!.root;
			const children = root.children;
			Object.defineProperty(root, 'children', { get: () => { reads++; return children; } });
		}
		queries(index, 'outside');
		queries(index, 'M1');
		reads = 0;
		for (let edit = 0; edit < 3; edit++) {
			index.setModule({ moduleName: 'm0', moduleKind: 'standard', source: source(`Changed${edit}`) });
			const fresh = new ProjectIndex();
			fresh.setModule({ moduleName: 'm0', moduleKind: 'standard', source: source(`Changed${edit}`) });
			for (let i = 1; i < count; i++) fresh.setModule({ moduleName: `M${i}`, moduleKind: 'standard', source: source(`N${i}`) });
			for (const caller of ['outside', 'M1']) expect(queries(index, caller)).toEqual(queries(fresh, caller));
		}
		expect(reads).toBe(0);
	});
}
it('invalidates both visibility sides on metadata, activity, removal and re-addition', () => {
	const index = new ProjectIndex();
	const input = { moduleName: 'Thing', moduleKind: 'class' as const, source: '#If VBA7 Then\nPublic Sub Active()\nEnd Sub\n#Else\nPrivate Sub Inactive()\nEnd Sub\n#End If\n' };
	index.setModule({ ...input, predeclaredId: true, conditionalCompilation: { compilerConstants: { VBA7: true } } });
	expect(index.visibleProcedureNames('Thing')).toContain('active');
	expect(index.visibleIdentifierNames('outside')).toContain('thing');
	queries(index, 'Thing');
	index.setModule({ ...input, moduleName: 'thing', predeclaredId: false, conditionalCompilation: { compilerConstants: { VBA7: false } } });
	const fresh = new ProjectIndex();
	fresh.setModule({ ...input, moduleName: 'thing', predeclaredId: false, conditionalCompilation: { compilerConstants: { VBA7: false } } });
	for (const caller of ['thing', 'outside']) expect(queries(index, caller)).toEqual(queries(fresh, caller));
	expect(index.visibleIdentifierNames('outside')).not.toContain('thing');
	expect(index.visibleProcedureNames('thing')).not.toContain('active');
	expect(index.visibleProcedureNames('thing')).toContain('inactive');
	index.removeModule('THING');
	expect(queries(index, 'outside')).toEqual(queries(new ProjectIndex(), 'outside'));
	index.setModule({ moduleName: 'THING', moduleKind: 'standard', source: source('Restored') });
	const restored = new ProjectIndex();
	restored.setModule({ moduleName: 'THING', moduleKind: 'standard', source: source('Restored') });
	for (const caller of ['thing', 'outside']) expect(queries(index, caller)).toEqual(queries(restored, caller));
});
