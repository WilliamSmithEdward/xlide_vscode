export interface VbaModuleDisplayInput {
    name?: string;
    moduleName?: string;
    type?: string;
    moduleType?: string;
}

const MODULE_TYPE_ORDER: Record<string, number> = {
    document: 0,
    userform: 1,
    accessform: 1,
    accessreport: 1,
    standard: 2,
    class: 3,
};

const MODULE_THEME_ICONS: Record<string, string> = {
    standard: 'symbol-module',
    class: 'symbol-class',
    document: 'symbol-namespace',
    userform: 'window',
    accessform: 'window',
    accessreport: 'book',
};

const MODULE_TYPE_BADGES: Record<string, string> = {
    document: 'D',
    userform: 'F',
    accessform: 'F',
    accessreport: 'R',
    standard: 'M',
    class: 'C',
};

export function compareVbaModulesForTreeOrder(
    left: VbaModuleDisplayInput,
    right: VbaModuleDisplayInput,
): number {
    const leftType = moduleTypeOf(left);
    const rightType = moduleTypeOf(right);
    const leftOrder = MODULE_TYPE_ORDER[leftType] ?? 4;
    const rightOrder = MODULE_TYPE_ORDER[rightType] ?? 4;
    if (leftOrder !== rightOrder) {
        return leftOrder - rightOrder;
    }
    return moduleNameOf(left).localeCompare(moduleNameOf(right));
}

export function moduleThemeIconName(moduleType: string | undefined): string {
    return MODULE_THEME_ICONS[normalizeModuleType(moduleType)] ?? 'symbol-module';
}

export function moduleTypeBadge(moduleType: string | undefined): string {
    return MODULE_TYPE_BADGES[normalizeModuleType(moduleType)] ?? 'M';
}

export function moduleTypeLabel(moduleType: string | undefined): string {
    const normalized = normalizeModuleType(moduleType);
    if (!normalized) {
        return 'module';
    }
    return normalized;
}

/** A module member as the explorer lists it: `Sub`, `Property Get`, `Enum`... */
export interface ModuleMemberEntry {
    name: string;
    kind: string;
    line: number;
}

export type ModuleMemberSort = 'code' | 'name' | 'kind';

const MEMBER_KIND_ORDER: Record<string, number> = { enum: 0, type: 1, sub: 2, function: 3, property: 4 };
const PROPERTY_PART_ORDER: Record<string, number> = { get: 0, let: 1, set: 2 };

/**
 * A module's members in the order the explorer lists them (issue #1341):
 * as written (`code`), by name, or by kind (Enums, Types, Subs, Functions,
 * then Properties), each kind by name. Under every order a property's Get,
 * Let and Set stay together, in that order, where the first of them would
 * sit. Members that tie keep the order they are written in.
 */
export function orderModuleMembers<T extends ModuleMemberEntry>(
    members: readonly T[],
    sort: ModuleMemberSort,
): T[] {
    const groups: { first: number; name: string; rank: number; members: T[] }[] = [];
    const properties = new Map<string, (typeof groups)[number]>();
    members.forEach((member, index) => {
        const word = member.kind.toLowerCase().split(/\s+/)[0];
        if (word === 'property') {
            const key = member.name.toLowerCase();
            const group = properties.get(key);
            if (group) {
                group.members.push(member);
                return;
            }
            const created = { first: index, name: member.name, rank: MEMBER_KIND_ORDER.property, members: [member] };
            properties.set(key, created);
            groups.push(created);
            return;
        }
        groups.push({ first: index, name: member.name, rank: MEMBER_KIND_ORDER[word] ?? MEMBER_KIND_ORDER.sub, members: [member] });
    });
    const partOf = (member: T): number => PROPERTY_PART_ORDER[member.kind.toLowerCase().split(/\s+/)[1] ?? ''] ?? 3;
    for (const group of properties.values()) {
        group.members.sort((left, right) => partOf(left) - partOf(right));
    }
    const byName = (left: (typeof groups)[number], right: (typeof groups)[number]): number =>
        left.name.localeCompare(right.name, undefined, { sensitivity: 'base' });
    groups.sort((left, right) => (sort === 'kind' ? left.rank - right.rank : 0)
        || (sort === 'code' ? 0 : byName(left, right))
        || left.first - right.first);
    return groups.flatMap((group) => group.members);
}

/** The icon of a member's row, from its label ("Enum Colors", "Sub Run"). */
export function memberThemeIconName(label: string): string {
    const word = label.split(' ', 1)[0];
    return word === 'Enum' ? 'symbol-enum' : word === 'Type' ? 'symbol-struct' : 'symbol-method';
}

function moduleTypeOf(input: VbaModuleDisplayInput): string {
    return normalizeModuleType(input.type ?? input.moduleType);
}

function moduleNameOf(input: VbaModuleDisplayInput): string {
    return input.name ?? input.moduleName ?? '';
}

function normalizeModuleType(moduleType: string | undefined): string {
    return (moduleType ?? '').trim().toLowerCase();
}
