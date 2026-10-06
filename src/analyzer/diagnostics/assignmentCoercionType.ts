import { isKnownScalarType, normalizeType } from './typeInference';
import type { MemberCompletionContext } from '../completion/memberAccess';
import { hostDisplayName, resolveHostEnum } from '../host/hostModel';
import { resolveVbaLibraryQualifier } from '../runtime/vbaRuntime';

/** Enum storage is Long, including unnamed numeric values. Cache within one analysis pass. */
export function createAssignmentCoercionType(ctx: MemberCompletionContext, sourceEnums: ReadonlySet<string> = new Set()): (declared: string) => string {
	const cache = new Map<string, string>();
	let project: Map<string, { enum: boolean; nonEnum: boolean }> | undefined;
	const projectTypes = () => {
		if (project) { return project; }
		project = new Map<string, { enum: boolean; nonEnum: boolean }>();
		for (const type of ctx.projectClassMembers ?? []) {
			for (const name of [type.name, `${type.moduleName}.${type.name}`]) {
				const key = name.toLowerCase();
				const entry = project.get(key) ?? { enum: false, nonEnum: false };
				entry.enum ||= type.kind === 'enum';
				entry.nonEnum ||= type.kind !== 'enum';
				project.set(key, entry);
			}
		}
		return project;
	};
	return declared => {
		if (isKnownScalarType(normalizeType(declared) ?? '')) { return declared; }
		const key = declared.trim().toLowerCase();
		const cached = cache.get(key);
		if (cached !== undefined) { return cached; }
		const surface = projectTypes().get(key);
		let result = declared;
		if (sourceEnums.has(key) || (surface?.enum && !surface.nonEnum)) {
			result = 'Long';
		} else if (!surface?.nonEnum) {
			const runtime = resolveVbaLibraryQualifier(declared.replace(/^VBA\./i, ''));
			const enumeration = resolveHostEnum(declared.split('.').at(-1)!, ctx.model);
			const prefix = declared.includes('.') ? declared.slice(0, declared.lastIndexOf('.')).toLowerCase() : undefined;
			if (runtime?.constants?.some(c => c.type === runtime.name)
				|| (enumeration && (!prefix || prefix === (enumeration.library ?? hostDisplayName(ctx.model)).toLowerCase()))) {
				result = 'Long';
			}
		}
		cache.set(key, result);
		return result;
	};
}
