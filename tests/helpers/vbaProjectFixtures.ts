import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectAnalysisOptions,
	type VbaProjectModuleInput,
} from '../../src/vbaProjectAnalysis';
import type {
	ProjectIndex,
	VbaProcedureSignature,
} from '../../src/analyzer';

export interface VbaProjectFixtureModule {
	name: string;
	type?: string;
	moduleKind?: VbaProjectModuleInput['moduleKind'];
	documentType?: VbaProjectModuleInput['documentType'];
	sourceLines: string[];
}

export interface VbaProjectFixtureTypeAssertion {
	name: string;
	kind: string;
	moduleName: string;
}

export interface VbaProjectFixtureMemberSurfaceAssertion {
	name: string;
	members: string[];
	absentMembers?: string[];
}

export interface VbaProjectFixtureModuleContextAssertion {
	moduleName: string;
	knownProcedures?: string[];
	absentKnownProcedures?: string[];
	projectTypes?: VbaProjectFixtureTypeAssertion[];
	absentProjectTypes?: string[];
	memberSurfaces?: VbaProjectFixtureMemberSurfaceAssertion[];
}

export interface VbaProjectFixtureMemberCompletionAssertion {
	moduleName: string;
	marker: string;
	include: string[];
	exclude?: string[];
}

export interface VbaProjectFixtureDiagnosticCodeAssertion {
	code: string;
	count: number;
	messagesContain?: string[];
}

export interface VbaProjectFixtureDiagnosticAssertion {
	moduleName: string;
	codes: VbaProjectFixtureDiagnosticCodeAssertion[];
}

export interface VbaProjectFixtureSemanticTokenAssertion {
	text: string;
	type: string;
}

export interface VbaProjectFixtureSemanticTokensAssertion {
	moduleName: string;
	tokens: VbaProjectFixtureSemanticTokenAssertion[];
}

export interface VbaProjectFixtureSignatureHelpAssertion {
	moduleName: string;
	marker: string;
	label: string;
	parameters?: string[];
	documentationContains?: string;
}

export interface VbaProjectFixtureHoverAssertion {
	moduleName: string;
	marker: string;
	signature: string;
	detailsContain?: string[];
	documentationContains?: string;
}

export interface VbaProjectFixtureAssertions {
	moduleContexts?: VbaProjectFixtureModuleContextAssertion[];
	memberCompletions?: VbaProjectFixtureMemberCompletionAssertion[];
	diagnostics?: VbaProjectFixtureDiagnosticAssertion[];
	semanticTokens?: VbaProjectFixtureSemanticTokensAssertion[];
	signatureHelp?: VbaProjectFixtureSignatureHelpAssertion[];
	hovers?: VbaProjectFixtureHoverAssertion[];
}

export interface VbaProjectFixture {
	version: 1;
	id: string;
	description: string;
	modules: VbaProjectFixtureModule[];
	assertions: VbaProjectFixtureAssertions;
}

export interface VbaProjectFixtureContext {
	project: ProjectIndex;
	options: VbaProjectAnalysisOptions;
	projectProcedures: VbaProcedureSignature[];
}

const fixtureRoot = join(process.cwd(), 'tests', 'fixtures', 'vbaProjects');

export function loadVbaProjectFixtures(): VbaProjectFixture[] {
	return readdirSync(fixtureRoot)
		.filter((name) => name.endsWith('.json'))
		.sort()
		.map((name) => readFixture(join(fixtureRoot, name)));
}

export function fixtureModules(fixture: VbaProjectFixture): VbaProjectModuleInput[] {
	return fixture.modules.map((mod) => ({
		moduleName: mod.name,
		type: mod.type,
		moduleKind: mod.moduleKind,
		documentType: mod.documentType,
		source: mod.sourceLines.join('\n'),
	}));
}

export function fixtureModule(
	fixture: VbaProjectFixture,
	moduleName: string,
): VbaProjectModuleInput {
	const mod = fixtureModules(fixture).find(
		(candidate) => candidate.moduleName.toLowerCase() === moduleName.toLowerCase(),
	);
	if (!mod) {
		throw new Error(`Fixture ${fixture.id} has no module named ${moduleName}`);
	}
	return mod;
}

export function buildFixtureProject(fixture: VbaProjectFixture): ProjectIndex {
	return buildVbaProjectIndex(fixtureModules(fixture));
}

export function fixtureContext(
	fixture: VbaProjectFixture,
	moduleName: string,
): VbaProjectFixtureContext {
	const project = buildFixtureProject(fixture);
	const procedureMap = projectProcedureSignatures(project);
	return {
		project,
		options: projectAnalysisOptionsForModule(project, moduleName, procedureMap),
		projectProcedures: uniqueProcedureList(procedureMap),
	};
}

function readFixture(path: string): VbaProjectFixture {
	const fixture = JSON.parse(readFileSync(path, 'utf8')) as VbaProjectFixture;
	validateFixture(fixture, path);
	return fixture;
}

function validateFixture(fixture: VbaProjectFixture, path: string): void {
	if (fixture.version !== 1) {
		throw new Error(`Unsupported VBA project fixture version in ${path}`);
	}
	if (!fixture.id || !fixture.description) {
		throw new Error(`VBA project fixture ${path} must declare id and description`);
	}
	if (!Array.isArray(fixture.modules) || fixture.modules.length === 0) {
		throw new Error(`VBA project fixture ${fixture.id} must include modules`);
	}
	const names = new Set<string>();
	for (const mod of fixture.modules) {
		if (!mod.name || !Array.isArray(mod.sourceLines)) {
			throw new Error(`VBA project fixture ${fixture.id} has an invalid module entry`);
		}
		const key = mod.name.toLowerCase();
		if (names.has(key)) {
			throw new Error(`VBA project fixture ${fixture.id} declares duplicate module ${mod.name}`);
		}
		names.add(key);
	}
}

function uniqueProcedureList(
	procedureMap: ReturnType<typeof projectProcedureSignatures>,
): VbaProcedureSignature[] {
	const out: VbaProcedureSignature[] = [];
	const seen = new Set<string>();
	for (const candidates of procedureMap?.values() ?? []) {
		for (const candidate of candidates) {
			const key = [
				candidate.moduleName.toLowerCase(),
				candidate.name.toLowerCase(),
				candidate.kind,
			].join(':');
			if (seen.has(key)) {
				continue;
			}
			seen.add(key);
			out.push(candidate);
		}
	}
	return out;
}
