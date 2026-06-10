import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import { xlideGlobalSettingManifest } from '../src/globalSettings';

function contributedConfigurationProperties(): Record<string, Record<string, unknown>> {
	const manifest = JSON.parse(readFileSync('package.json', 'utf8')) as {
		contributes?: { configuration?: { properties?: Record<string, Record<string, unknown>> } };
	};
	return manifest.contributes?.configuration?.properties ?? {};
}

describe('global settings manifest', () => {
	it('keeps package.json contributes.configuration in sync with the schema registry', () => {
		const properties = contributedConfigurationProperties();
		const expected = xlideGlobalSettingManifest();

		expect(Object.keys(properties).sort()).toEqual(Object.keys(expected).sort());
		for (const [key, schema] of Object.entries(expected)) {
			const { description, ...contributed } = properties[key];
			expect(contributed, `contributed schema for ${key}`).toEqual(schema);
			expect(description, `description for ${key}`).toMatch(/\S/);
		}
	});
});
