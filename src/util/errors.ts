export function errorMessage(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
}

/** True for the errors Node's file and process APIs throw, which carry a `code`. */
export function isNodeError(value: unknown): value is NodeJS.ErrnoException {
    return value !== null && typeof value === 'object' && 'code' in value;
}
