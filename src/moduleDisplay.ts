export interface VbaModuleDisplayInput {
    name?: string;
    moduleName?: string;
    type?: string;
    moduleType?: string;
}

const MODULE_TYPE_ORDER: Record<string, number> = {
    document: 0,
    userform: 1,
    standard: 2,
    class: 3,
};

const MODULE_THEME_ICONS: Record<string, string> = {
    standard: 'symbol-module',
    class: 'symbol-class',
    document: 'symbol-namespace',
    userform: 'window',
};

const MODULE_TYPE_BADGES: Record<string, string> = {
    document: 'D',
    userform: 'F',
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

function moduleTypeOf(input: VbaModuleDisplayInput): string {
    return normalizeModuleType(input.type ?? input.moduleType);
}

function moduleNameOf(input: VbaModuleDisplayInput): string {
    return input.name ?? input.moduleName ?? '';
}

function normalizeModuleType(moduleType: string | undefined): string {
    return (moduleType ?? '').trim().toLowerCase();
}
