import type { HostEnum } from './excelObjectModel';

/** Preserve the defining library when a host also exposes shared reference enums. */
export function enumsFromLibrary(enums: Readonly<Record<string, HostEnum>>, library: string): Record<string, HostEnum> {
	return Object.fromEntries(Object.entries(enums).map(([key, value]) => [key, { ...value, library: value.library ?? library }]));
}
