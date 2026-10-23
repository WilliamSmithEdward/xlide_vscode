// An oracle fixture's module as the analyzer's project input, with what its
// designer and attributes declare: a UserForm's controls, a document
// module's designer class, a class's PredeclaredId (issue #1419). The VBE
// oracle builds the same module from the same fields, so the analyzer sees
// the project the VBE compiled.

import type { EventHandlerDocumentType } from '../../src/analyzer';
import { hostTokenForLibid } from '../../src/analyzer/host/hostLibraries';
import type { VbaProjectModuleInput } from '../../src/vbaProjectAnalysis';

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
		source: mod.source,
		...(implicitMembers !== undefined ? { implicitMembers } : {}),
		...(mod.designerClass !== undefined ? { designerClass: mod.designerClass } : {}),
		...(mod.predeclaredId !== undefined ? { predeclaredId: mod.predeclaredId } : {}),
	};
}
