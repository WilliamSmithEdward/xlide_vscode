import type { MemberCompletionContext } from './memberAccess';

const exportedSetterTypes = new WeakMap<NonNullable<MemberCompletionContext['projectClassMembers']>, Map<string, string | undefined>>();
export function projectSetterValueType(ctx: MemberCompletionContext, moduleName: string, name: string): string | undefined {
	const surfaces = ctx.projectClassMembers;
	if (!surfaces) { return undefined; }
	let types = exportedSetterTypes.get(surfaces);
	if (!types) {
		types = new Map();
		for (const surface of surfaces) {
			if (surface.kind !== 'standardModule') { continue; }
			for (const member of surface.members) {
				if (member.letAccessor) { types.set(`${surface.moduleName}.${member.name}`.toLowerCase(), member.procedureParams?.propertyLet?.at(-1)?.type ?? member.writeType); }
			}
		}
		exportedSetterTypes.set(surfaces, types);
	}
	return types.get(`${moduleName}.${name}`.toLowerCase());
}

