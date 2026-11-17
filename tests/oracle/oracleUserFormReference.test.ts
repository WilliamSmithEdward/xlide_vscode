import { describe, expect, it } from 'vitest';
import { oracleProjectReferences } from '../helpers/oracleModuleInput';

// Adding a UserForm adds the MSForms reference, in Excel, Word, PowerPoint
// and Access 16.0 (build 20527, measured), so the file the oracle compiled a
// fixture with a UserForm in records MSForms, and the extension reads it.
describe('the references of a fixture with a UserForm', () => {
	const form = { name: 'Form1', type: 'userform', source: '' };
	const module = { name: 'Module1', type: 'standard', source: '' };

	it.each([
		['excel', ['stdole', 'Office']],
		['word', ['stdole', 'Normal', 'Office']],
		['powerpoint', ['stdole', 'Office']],
		['access', ['stdole', 'DAO']],
	])('%s adds MSForms after its defaults, and only with a UserForm', (host, defaults) => {
		expect(oracleProjectReferences({ host, modules: [module, form] }).referencedLibraries).toEqual([...defaults, 'MSForms']);
		expect(oracleProjectReferences({ host, modules: [module] }).referencedLibraries).toEqual(defaults);
		expect(oracleProjectReferences({ host }).referencedLibraries).toEqual(defaults);
	});

	it('MSForms comes after the references the fixture adds, as the oracle adds the form after them', () => {
		const scripting = { guid: '{420B2830-E718-11CF-893D-00A0C9054228}' };
		expect(oracleProjectReferences({ host: 'excel', modules: [form], references: [scripting] }).referencedLibraries)
			.toEqual(['stdole', 'Office', 'Scripting', 'MSForms']);
	});
});
