import { describe, expect, it } from 'vitest';
import {
	analyzeModule,
	resolveHover,
	resolveIdentifierCompletions,
	resolveMemberCompletions,
	resolveSignatureHelp,
	resolveTypeCompletions,
	resolveTypeSemanticTokens,
} from '../src/analyzer';
import {
	fixtureContext,
	fixtureModule,
	loadVbaProjectFixtures,
	type VbaProjectFixture,
} from './helpers/vbaProjectFixtures';

function offsetAfterMarker(source: string, marker: string): number {
	const idx = source.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function offsetInsideMarker(source: string, marker: string): number {
	const idx = source.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + Math.min(1, Math.max(marker.length - 1, 0));
}

function editorContext(fixture: VbaProjectFixture, moduleName: string) {
	const { options, projectProcedures, projectSymbols } = fixtureContext(fixture, moduleName);
	const mod = fixture.modules.find(
		(candidate) => candidate.name.toLowerCase() === moduleName.toLowerCase(),
	);
	return {
		moduleName,
		projectClassMembers: options.projectClassMembers,
		projectTypes: options.projectTypes,
		projectProcedures,
		projectSymbols,
		implicitMembers: options.implicitMembers,
		// What `Me` denotes, as the editor's own context service derives it.
		meProjectType: mod && ['class', 'document', 'userform'].includes(mod.type ?? '')
			? mod.name
			: undefined,
		meType: mod?.type === 'userform' ? 'MSForms.UserForm' : undefined,
	};
}

function fixtureCodeNames(fixture: VbaProjectFixture): string[] {
	return fixture.modules
		.filter((mod) => mod.type === 'document')
		.map((mod) => mod.name);
}

describe('machine-readable VBA workbook project fixtures', () => {
	for (const fixture of loadVbaProjectFixtures()) {
		describe(fixture.id, () => {
			it('declares a valid workbook-level fixture shape', () => {
				expect(fixture.version).toBe(1);
				expect(fixture.modules.length).toBeGreaterThan(0);
				expect(fixture.assertions).toBeDefined();
			});

			for (const assertion of fixture.assertions.moduleContexts ?? []) {
				it(`derives project context for ${assertion.moduleName}`, () => {
					const { options } = fixtureContext(fixture, assertion.moduleName);
					const typeKeys = new Set(
						(options.projectTypes ?? []).map((type) =>
							`${type.name.toLowerCase()}:${type.kind}:${type.moduleName?.toLowerCase() ?? ''}`),
					);
					const knownProcedures = options.knownProcedures ?? new Set<string>();

					for (const procedure of assertion.knownProcedures ?? []) {
						expect(knownProcedures.has(procedure.toLowerCase()), procedure).toBe(true);
					}
					for (const procedure of assertion.absentKnownProcedures ?? []) {
						expect(knownProcedures.has(procedure.toLowerCase()), procedure).toBe(false);
					}
					for (const type of assertion.projectTypes ?? []) {
						expect(
							typeKeys.has(`${type.name.toLowerCase()}:${type.kind}:${type.moduleName.toLowerCase()}`),
							`${type.kind} ${type.name}`,
						).toBe(true);
					}
					for (const name of assertion.absentProjectTypes ?? []) {
						expect(
							(options.projectTypes ?? []).some((type) => type.name.toLowerCase() === name.toLowerCase()),
							name,
						).toBe(false);
					}
					for (const surfaceAssertion of assertion.memberSurfaces ?? []) {
						const surface = options.projectClassMembers?.find(
							(candidate) =>
								candidate.name.toLowerCase() === surfaceAssertion.name.toLowerCase(),
						);
						expect(surface, surfaceAssertion.name).toBeDefined();
						const members = new Set(surface?.members.map((member) => member.name.toLowerCase()));
						for (const member of surfaceAssertion.members) {
							expect(members.has(member.toLowerCase()), member).toBe(true);
						}
						for (const member of surfaceAssertion.absentMembers ?? []) {
							expect(members.has(member.toLowerCase()), member).toBe(false);
						}
					}
				});
			}

			for (const assertion of fixture.assertions.memberCompletions ?? []) {
				it(`resolves member completion at ${assertion.moduleName}:${assertion.marker}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const ctx = editorContext(fixture, assertion.moduleName);
					const names = resolveMemberCompletions(
						mod.source,
						offsetAfterMarker(mod.source, assertion.marker),
						ctx,
					).map((member) => member.name);

					for (const name of assertion.include) {
						expect(names, name).toContain(name);
					}
					for (const name of assertion.exclude ?? []) {
						expect(names, name).not.toContain(name);
					}
				});
			}

			for (const assertion of fixture.assertions.typeCompletions ?? []) {
				it(`resolves type completion at ${assertion.moduleName}:${assertion.marker}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const ctx = editorContext(fixture, assertion.moduleName);
					const names = resolveTypeCompletions(
						mod.source,
						offsetAfterMarker(mod.source, assertion.marker),
						ctx,
					).map((type) => type.name);

					for (const name of assertion.include) {
						expect(names, name).toContain(name);
					}
					for (const name of assertion.exclude ?? []) {
						expect(names, name).not.toContain(name);
					}
				});
			}

			for (const assertion of fixture.assertions.identifierCompletions ?? []) {
				it(`resolves identifier completion at ${assertion.moduleName}:${assertion.marker}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const ctx = {
						...editorContext(fixture, assertion.moduleName),
						codeNames: fixtureCodeNames(fixture),
					};
					const names = resolveIdentifierCompletions(
						mod.source,
						offsetAfterMarker(mod.source, assertion.marker),
						ctx,
					).map((identifier) => identifier.name);

					for (const name of assertion.include) {
						expect(names, name).toContain(name);
					}
					for (const name of assertion.exclude ?? []) {
						expect(names, name).not.toContain(name);
					}
				});
			}

			for (const assertion of fixture.assertions.diagnostics ?? []) {
				it(`matches diagnostics for ${assertion.moduleName}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const { options } = fixtureContext(fixture, assertion.moduleName);
					const diagnostics = analyzeModule(mod.source, {
						moduleName: assertion.moduleName,
						...options,
					});

					for (const codeAssertion of assertion.codes) {
						const hits = diagnostics.filter((diagnostic) => diagnostic.code === codeAssertion.code);
						expect(hits, codeAssertion.code).toHaveLength(codeAssertion.count);
						const messages = hits.map((hit) => hit.message).join('\n');
						for (const text of codeAssertion.messagesContain ?? []) {
							expect(messages, text).toContain(text);
						}
					}
				});
			}

			for (const assertion of fixture.assertions.semanticTokens ?? []) {
				it(`matches semantic type tokens for ${assertion.moduleName}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const { options } = fixtureContext(fixture, assertion.moduleName);
					const tokens = resolveTypeSemanticTokens(mod.source, {
						projectTypes: options.projectTypes,
					}).map((token) => ({
						text: mod.source.slice(token.span.start, token.span.end),
						type: token.tokenType,
					}));

					expect(tokens).toEqual(assertion.tokens);
				});
			}

			for (const assertion of fixture.assertions.signatureHelp ?? []) {
				it(`matches signature help for ${assertion.moduleName}:${assertion.marker}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const info = resolveSignatureHelp(
						mod.source,
						offsetAfterMarker(mod.source, assertion.marker),
						editorContext(fixture, assertion.moduleName),
					);

					expect(info?.label).toBe(assertion.label);
					if (assertion.parameters) {
						expect(info?.parameters.map((param) => param.label)).toEqual(assertion.parameters);
					}
					if (assertion.documentationContains) {
						expect(info?.documentation).toContain(assertion.documentationContains);
					}
				});
			}

			for (const assertion of fixture.assertions.hovers ?? []) {
				it(`matches hover for ${assertion.moduleName}:${assertion.marker}`, () => {
					const mod = fixtureModule(fixture, assertion.moduleName);
					const info = resolveHover(
						mod.source,
						offsetInsideMarker(mod.source, assertion.marker),
						editorContext(fixture, assertion.moduleName),
					);

					expect(info?.signature).toBe(assertion.signature);
					for (const detail of assertion.detailsContain ?? []) {
						expect(info?.details).toContain(detail);
					}
					if (assertion.documentationContains) {
						expect(info?.documentation).toContain(assertion.documentationContains);
					}
				});
			}
		});
	}
});
