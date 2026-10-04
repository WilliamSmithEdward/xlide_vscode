import {beforeEach, expect, it, vi} from 'vitest';
import type {HostObjectModel} from '../src/analyzer/host/excelObjectModel';

beforeEach(() => vi.resetModules());
function model(version: string): HostObjectModel {
	return {source: version, hostName: version, globalType: version + '.Global',
		types: {[version + '.Type']: {displayName: version, members: [{name: version + 'Member', kind: 'method'}]}},
		aliases: {shared: version + '.Type'}, globals: {Shared: version + '.Type'},
		constants: {Shared: {name: 'Shared', value: version}}, enums: {Shared: {displayName: version}},
		memberSignatures: {[version + '.Type']: {member: version + '()'}},
	};
}
function merged(first: HostObjectModel, second: HostObjectModel): HostObjectModel {
	return {source: first.source + ' + ' + second.source, hostName: first.hostName, globalType: first.globalType,
		types: {...second.types, ...first.types}, aliases: {...second.aliases, ...first.aliases}, globals: {...second.globals, ...first.globals},
		constants: {...second.constants, ...first.constants}, enums: {...Object.fromEntries(Object.entries(second.enums ?? {}).map(([key, value]) => [key, {...value, library: second.hostName}])), ...first.enums},
		memberSignatures: {...second.memberSignatures, ...first.memberSignatures},
	};
}

it.each([0, 1])('refreshes complete merged metadata when provider %i is replaced', async position => {
	const registry = await import('../src/analyzer/host/hostRegistry');
	const models = [model('Primary'), model('Secondary')];
	registry.registerHostObjectModel('project', () => models[0]);
	registry.registerHostObjectModel('visio', () => models[1]);
	const previous = registry.hostObjectModelForTokens(['project', 'visio']);
	const snapshot = structuredClone(previous);
	expect(previous).toEqual(merged(models[0], models[1]));
	const next = model('Replacement');
	registry.registerHostObjectModel(position === 0 ? 'project' : 'visio', () => next);
	models[position] = next;
	const actual = registry.hostObjectModelForTokens(['project', 'visio']);
	expect(actual).toEqual(merged(models[0], models[1]));
	expect(actual).not.toBe(previous);
	expect(registry.hostObjectModelForTokens(['project', 'visio'])).toBe(actual);
	expect(previous).toEqual(snapshot);
});

it('invalidates every cached ordering that includes the replaced token', async () => {
	const registry = await import('../src/analyzer/host/hostRegistry');
	const old = model('Old'), other = model('Other'), next = model('Next');
	registry.registerHostObjectModel('project', () => old);
	registry.registerHostObjectModel('visio', () => other);
	for (const tokens of [['project', 'visio'], ['visio', 'project'], ['project', 'project', 'visio']]) { registry.hostObjectModelForTokens(tokens); }
	registry.registerHostObjectModel('project', () => next);
	expect(registry.hostObjectModelForTokens(['project', 'visio'])).toEqual(merged(next, other));
	expect(registry.hostObjectModelForTokens(['visio', 'project'])).toEqual(merged(other, next));
	expect(registry.hostObjectModelForTokens(['project', 'project', 'visio'])?.source).toBe('Next + Next + Other');
});

it('treats re-registration of the same factory as a refresh', async () => {
	const registry = await import('../src/analyzer/host/hostRegistry');
	let current = model('Old');
	const provider = () => current;
	registry.registerHostObjectModel('project', provider);
	registry.registerHostObjectModel('visio', () => model('Other'));
	registry.hostObjectModelForTokens(['project', 'visio']);
	current = model('New');
	registry.registerHostObjectModel('project', provider);
	expect(registry.hostObjectModelForTokens(['project', 'visio'])?.source).toBe('New + Other');
});

it('keeps unaffected merged models cached', async () => {
	const registry = await import('../src/analyzer/host/hostRegistry');
	registry.registerHostObjectModel('visio', () => model('Visio'));
	registry.registerHostObjectModel('outlook', () => model('Outlook'));
	const before = registry.hostObjectModelForTokens(['visio', 'outlook']);
	registry.registerHostObjectModel('project', () => model('Project'));
	expect(registry.hostObjectModelForTokens(['visio', 'outlook'])).toBe(before);
});

it('includes newly registered tokens and retains absent/excel defaults', async () => {
	const registry = await import('../src/analyzer/host/hostRegistry');
	registry.registerHostObjectModel('visio', () => model('Visio'));
	registry.registerHostObjectModel('outlook', () => model('Outlook'));
	registry.hostObjectModelForTokens(['project', 'visio', 'outlook']);
	registry.registerHostObjectModel('project', () => model('Project'));
	expect(registry.hostObjectModelForTokens(['project', 'visio', 'outlook'])?.source).toBe('Project + Visio + Outlook');
	expect(registry.hostObjectModelForTokens([])).toBeUndefined();
	expect(registry.hostObjectModelForTokens(['excel'])).toBeUndefined();
	expect(registry.hostObjectModelForToken('project')?.hostName).toBe('Project');
});
