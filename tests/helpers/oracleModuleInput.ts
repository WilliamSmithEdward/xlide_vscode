import type { VbaProjectModuleInput } from '../../src/vbaProjectAnalysis';

/** Module metadata shared by the oracle's project-aware analyzer consumers. */
export interface OracleModule {
	name: string;
	type?: string;
	source: string;
	entry?: boolean;
	moduleKind?: VbaProjectModuleInput['moduleKind'];
	implicitMembers?: VbaProjectModuleInput['implicitMembers'];
	predeclaredId?: boolean;
}

export function oracleModuleInput(mod: OracleModule): VbaProjectModuleInput {
	return {
		moduleName: mod.name,
		type: mod.type,
		source: mod.source,
		moduleKind: mod.moduleKind,
		implicitMembers: mod.implicitMembers,
		predeclaredId: mod.predeclaredId,
	};
}
