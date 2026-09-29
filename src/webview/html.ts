import * as crypto from 'crypto';

export function escapeHtml(value: unknown): string {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/** Alias of escapeHtml: the escaping covers attribute position (double or single quoted) too. */
export function escapeAttr(value: unknown): string {
    return escapeHtml(value);
}

export function scriptJson(value: unknown): string {
    return JSON.stringify(value).replace(/</g, '\\u003c');
}

/** 128 random bits as hex: every value equally likely, unlike a byte taken modulo an alphabet. */
export function randomNonce(): string {
    return crypto.randomBytes(16).toString('hex');
}
