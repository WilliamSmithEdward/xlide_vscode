// An oracle fixture's module as the analyzer's project input, with what its
// designer and attributes declare: a UserForm's controls, a document
// module's designer class, a class's PredeclaredId (issue #1419). The VBE
// oracle builds the same module from the same fields, so the analyzer sees
// the project the VBE compiled.

import type { EventHandlerDocumentType } from '../../src/analyzer';
import { hostTokenForLibid } from '../../src/analyzer/host/hostLibraries';
import type { VbaProjectModuleInput } from '../../src/vbaProjectAnalysis';
import { referencedProjectSurface, type ReferencedProjectSurface } from '../../src/analyzer/symbols/referencedProjects';

export interface OracleFixtureModule {
	name: string;
	type?: string;
	source: string;
	/** Controls the oracle adds to a UserForm's designer, by ProgID. */
	designerControls?: readonly { name: string; progId: string }[];
	/** Members the designer declares, where no designerControls name them (a worksheet's ActiveX controls). */
	implicitMembers?: readonly { name: string; type: string }[];
	designerClass?: string;
	predeclaredId?: boolean;
	/** The code name the oracle gives a document module, `ThisWorkbook` renamed `BookCode`. */
	codeName?: string;
}

/** `Forms.CommandButton.1` is an `MSForms.CommandButton`. */
function controlType(progId: string): string {
	const match = /^Forms\.(\w+)\.\d+$/.exec(progId);
	if (!match) {
		throw new Error(`No MSForms type for the ProgID ${progId}`);
	}
	return `MSForms.${match[1]}`;
}

/**
 * The libraries a fixture's `references` add, as the analysis option the
 * extension derives from a project's references: a case the oracle compiled
 * with Outlook referenced is analyzed against Outlook's model too.
 */
export function oracleReferencedHosts(kase: { references?: readonly { guid: string }[] }): { referencedHosts?: string[] } {
	const tokens = (kase.references ?? []).flatMap((reference) => {
		const token = hostTokenForLibid(reference.guid);
		return token === undefined ? [] : [token];
	});
	return tokens.length === 0 ? {} : { referencedHosts: tokens };
}

/** The name the dir stream records for each library the corpus adds (measured, issue #1413). */
const RECORDED_LIBRARY_NAMES: Readonly<Record<string, string>> = {
	'{420B2830-E718-11CF-893D-00A0C9054228}': 'Scripting',
	'{3F4DACA7-160D-11D2-A8E9-00104B365C9F}': 'VBScript_RegExp_55',
	'{B691E011-1797-432E-907A-4D8C69339129}': 'ADODB',
	'{F5078F18-C551-11D3-89B9-0000F81FE221}': 'MSXML2',
	'{0002E157-0000-0000-C000-000000000046}': 'VBIDE',
	'{662901FC-6951-4854-9EB2-D9A2570F2B2E}': 'WinHttp',
	'{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}': 'Shell32',
	'{00062FFF-0000-0000-C000-000000000046}': 'Outlook',
	'{00020813-0000-0000-C000-000000000046}': 'Excel',
	'{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}': 'DAO',
};
/** What a new file of each host records before any reference is added, and its project's name (measured, issue #1413). */
const DEFAULT_REFERENCES: Readonly<Record<string, readonly string[]>> = {
	excel: ['stdole', 'Office'], word: ['stdole', 'Normal', 'Office'], powerpoint: ['stdole', 'Office'], access: ['stdole', 'DAO'],
};
const DEFAULT_PROJECT_NAMES: Readonly<Record<string, string>> = { excel: 'VBAProject', word: 'Project', powerpoint: 'VBAProject', access: 'Database1' };

/**
 * The references and project name the extension reads from the file the
 * oracle compiled the fixture in: the host's defaults, then what the
 * fixture's `references` add.
 */
export function oracleProjectReferences(kase: {
	host?: string;
	references?: readonly { guid: string }[];
	projectReference?: { name: string; modules: readonly OracleFixtureModule[] };
}): {
	referencedLibraries: string[];
	projectName: string;
	referencedProjects?: string[];
	referencedProjectSurfaces?: ReferencedProjectSurface[];
} {
	const host = kase.host ?? 'excel';
	const added = (kase.references ?? []).map((reference) => {
		const name = RECORDED_LIBRARY_NAMES[reference.guid];
		if (!name) {
			throw new Error(`No recorded name for the reference ${reference.guid}; measure it and add it here`);
		}
		return name;
	});
	// A VBA project the oracle referenced, read as the engine reads one (issue #1503).
	const project = kase.projectReference;
	return {
		referencedLibraries: [...new Set([...DEFAULT_REFERENCES[host], ...added, ...(project ? [project.name] : [])])],
		projectName: DEFAULT_PROJECT_NAMES[host],
		...(project ? {
			referencedProjects: [project.name],
			referencedProjectSurfaces: [referencedProjectSurface(project.name, project.modules.map((mod) => ({
				name: mod.name,
				kind: mod.type === 'class' ? 'class' as const : 'standard' as const,
				source: mod.source,
			})))],
		} : {}),
	};
}

/** The document modules a new file of each host has before the fixture's own modules are added. */
const DEFAULT_DOCUMENT_MODULES: Readonly<Record<string, readonly string[]>> = {
	excel: ['ThisWorkbook', 'Sheet1'], word: ['ThisDocument'], powerpoint: [], access: [],
};

/**
 * The host's own document modules a fixture leaves out, as empty modules: the
 * file the oracle compiled it in has them, so `Dim a As ThisWorkbook` names
 * a type of the project.
 */
export function oracleDefaultDocumentModules(kase: { host?: string }, modules: readonly OracleFixtureModule[]): OracleFixtureModule[] {
	const named = new Set(modules.map((mod) => (mod.codeName ?? mod.name).toLowerCase()));
	return DEFAULT_DOCUMENT_MODULES[kase.host ?? 'excel']
		.filter((name) => !named.has(name.toLowerCase()) && !modules.some((mod) => mod.name.toLowerCase() === name.toLowerCase()))
		.map((name) => ({ name, type: 'document', source: '' }));
}

/** What a new file's document module is, by the name it was made with. */
function documentTypeOf(name: string): EventHandlerDocumentType {
	switch (name.toLowerCase()) {
		case 'thisworkbook': return 'workbook';
		case 'thisdocument': return 'document';
		default: return 'worksheet';
	}
}

export function oracleModuleInput(mod: OracleFixtureModule): VbaProjectModuleInput {
	const implicitMembers = mod.designerControls
		? mod.designerControls.map((control) => ({ name: control.name, type: controlType(control.progId) }))
		: mod.implicitMembers;
	return {
		moduleName: mod.codeName ?? mod.name,
		// A renamed module is still the object it was made as, which the
		// extension reads from the module's VB_Base.
		...(mod.codeName !== undefined ? { documentType: documentTypeOf(mod.name) } : {}),
		type: mod.type,
		// The engine lists an Access form or report as a designer module
		// (accessform), as the oracle's document modules of Access are.
		...(mod.designerClass?.startsWith('Access.') ? { moduleKind: 'userform' as const } : {}),
		source: mod.source,
		...(implicitMembers !== undefined ? { implicitMembers } : {}),
		...(mod.designerClass !== undefined ? { designerClass: mod.designerClass } : {}),
		...(mod.predeclaredId !== undefined ? { predeclaredId: mod.predeclaredId } : {}),
	};
}
