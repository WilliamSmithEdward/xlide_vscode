import {afterEach, expect, it, vi} from 'vitest';
import {ProjectIndex} from '../src/analyzer/symbols/projectIndex';
import * as facts from '../src/analyzer/symbols/classMemberFacts';

afterEach(() => vi.restoreAllMocks());
function source(eol: string, assigned = false) {
	return ['Public Ref As Object', 'Public EmptyValue As Variant', 'Public Function ScalarValue() As Variant', 'ScalarValue = 42', 'End Function', ...(assigned ? ['Private Sub AssignIt()', 'Set Ref = New Collection', 'End Sub'] : [])].join(eol);
}
function project(count: number, eol: string) {
	const index = new ProjectIndex();
	for (let i = 0; i < count; i++) { index.setModule({moduleName: 'Class' + i, moduleKind: 'class', source: source(eol)}); }
	index.setModule({moduleName: 'Caller', moduleKind: 'standard', source: 'Sub S()' + eol + 'End Sub'});
	return index;
}

it.each([10, 1000].flatMap(count => ['\n', '\r\n', '\r'].map(eol => ({count, eol}))))('reuses $count unchanged class value scans after a caller edit with EOL $eol', ({count, eol}) => {
	const index = project(count, eol), before = structuredClone(index.projectClassMembers());
	const scan = vi.spyOn(facts, 'classMemberValues');
	index.setModule({moduleName: 'Caller', moduleKind: 'standard', source: 'Sub S()' + eol + 'Dim n As Long' + eol + 'End Sub'});
	const after = index.projectClassMembers();
	expect(after).toEqual(before);
	for (const type of after) { expect(type.members.map(member => [member.name, member.knownValue])).toEqual([['Ref', 'nothing'], ['EmptyValue', 'empty'], ['ScalarValue', 'scalar']]); }
	expect(scan).not.toHaveBeenCalled();
});

it('recomputes only an edited class and keeps editor snapshots free of value annotations', () => {
	const index = project(10, '\n');
	index.projectClassMembers();
	const scan = vi.spyOn(facts, 'classMemberValues');
	index.setModule({moduleName: 'Class4', moduleKind: 'class', source: source('\n', true)});
	const editor = index.projectClassMembers({includeClassValueFacts: false});
	expect(scan).not.toHaveBeenCalled();
	const snapshot = structuredClone(editor), actual = index.projectClassMembers();
	expect(scan).toHaveBeenCalledTimes(1);
	expect(scan.mock.calls[0][0]).toBe(source('\n', true));
	expect(actual.find(type => type.name === 'Class4')!.members[0].knownValue).toBeUndefined();
	expect(editor).toEqual(snapshot);
	const fresh = project(10, '\n');
	fresh.setModule({moduleName: 'Class4', moduleKind: 'class', source: source('\n', true)});
	expect(actual).toEqual(fresh.projectClassMembers());
});

it('drops removed and replaced class contributions and preserves old snapshots', () => {
	const index = project(1, '\n'), original = index.projectClassMembers(), copy = structuredClone(original);
	index.removeModule('CLASS0');
	expect(index.projectClassMembers()).toEqual([]);
	index.setModule({moduleName: 'Class0', moduleKind: 'class', source: source('\n', true)});
	expect(index.projectClassMembers()[0].members[0].knownValue).toBeUndefined();
	expect(original).toEqual(copy);
	index.setModule({moduleName: 'Class0', moduleKind: 'document', source: source('\n')});
	expect(index.projectClassMembers()[0].members.every(member => member.knownValue === undefined)).toBe(true);
	index.setModule({moduleName: 'Class0', moduleKind: 'class', source: source('\n')});
	expect(index.projectClassMembers()[0].members[0].knownValue).toBe('nothing');
});
