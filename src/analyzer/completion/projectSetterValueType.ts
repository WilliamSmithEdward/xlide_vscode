import type { MemberCompletionContext } from './memberAccess';
import type { VbaProjectClassMember } from '../symbols/symbolModel';

const standardMembers = new WeakMap<NonNullable<MemberCompletionContext['projectClassMembers']>, Map<string, VbaProjectClassMember>>();
function member(ctx: MemberCompletionContext, moduleName: string, name: string): VbaProjectClassMember | undefined {
	const surfaces = ctx.projectClassMembers;
	if (!surfaces) { return undefined; }
	let members = standardMembers.get(surfaces);
	if (!members) {
		members = new Map();
		for (const surface of surfaces) {
			if (surface.kind !== 'standardModule') { continue; }
			for (const member of surface.members) { members.set(`${surface.moduleName}.${member.name}`.toLowerCase(), member); }
		}
		standardMembers.set(surfaces, members);
	}
	return members.get(`${moduleName}.${name}`.toLowerCase());
}

export function projectSetterValueType(ctx: MemberCompletionContext, moduleName: string, name: string): string | undefined {
	const setter = member(ctx, moduleName, name);
	return setter?.letAccessor ? setter.procedureParams?.propertyLet?.at(-1)?.type ?? setter.writeType : undefined;
}

export function projectGetterKnownValue(ctx: MemberCompletionContext, moduleName: string, name: string): VbaProjectClassMember['knownValue'] {
	return member(ctx, moduleName, name)?.knownValue;
}

export function projectSetterParameters(ctx: MemberCompletionContext, moduleName: string, name: string, kind: 'propertyLet' | 'propertySet') {
	return member(ctx,moduleName,name)?.procedureParams?.[kind];
}
