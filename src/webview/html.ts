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

export function randomNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    const bytes = crypto.randomBytes(32);
    let nonce = '';
    for (let i = 0; i < bytes.length; i++) {
        nonce += chars.charAt(bytes[i] % chars.length);
    }
    return nonce;
}
