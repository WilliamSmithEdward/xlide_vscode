// Existing Office project password verification, [MS-OVBA] 2.3.1.15-17,
// 2.4.3.3 and 2.4.4. This never removes or rewrites project protection.
import { projectIdentityKey } from '../projectIdentity';
import { createHash } from 'crypto';
import { decodeCodePage, encodeCodePage } from './codePages';

const INVALID = 'The VBA project protection is malformed or unsupported. Unlock and save the project in its Office application before using XLIDE.';

export class VbaProjectLockedError extends Error {
    constructor() { super('This VBA project is password-protected. Enter its password in XLIDE to access the file.'); }
}

/** Strict decoding: damaged records must never turn into an unprotected project. */
export function decodeProtectionRecord(hex: string): Buffer {
    if (!/^(?:[0-9a-f]{2})+$/i.test(hex)) { throw new Error(INVALID); }
    const bytes = Buffer.from(hex, 'hex');
    if (bytes.length < 8 || (bytes[0] ^ bytes[1]) !== 2) { throw new Error(INVALID); }
    const ignored = (bytes[0] & 6) >> 1;
    if (bytes.length < 7 + ignored) { throw new Error(INVALID); }
    let plain = bytes[0] ^ bytes[2];
    let previous = bytes[2];
    let older = bytes[1];
    const decoded = Buffer.alloc(bytes.length - 3);
    for (let i = 3; i < bytes.length; i++) {
        const value = bytes[i] ^ ((older + plain) & 255);
        older = previous;
        previous = bytes[i];
        plain = value;
        decoded[i - 3] = value;
    }
    const length = decoded.readUInt32LE(ignored);
    const data = decoded.subarray(ignored + 4);
    if (length !== data.length) { throw new Error(INVALID); }
    return data;
}

export class VbaProjectProtection {
    readonly hasPassword: boolean;
    readonly requiresPassword: boolean;
    readonly fingerprint: string;
    private readonly passwordData?: Buffer;
    private readonly invalid: boolean;

    constructor(raw: Buffer | undefined, private readonly codePage: number) {
        const records = new Map<string, string>();
        let invalid = raw === undefined;
        for (const line of decodeCodePage(raw ?? Buffer.alloc(0), codePage).split(/\r?\n/)) {
            if (!/^\s*(CMG|DPB|GC)\s*=/i.test(line)) { continue; }
            const match = /^\s*(CMG|DPB|GC)\s*=\s*"([0-9a-f]+)"\s*$/i.exec(line);
            if (!match || records.has(match[1].toUpperCase())) { invalid = true; continue; }
            records.set(match[1].toUpperCase(), match[2].toUpperCase());
        }
        let data: Buffer | undefined;
        let locked = false;
        try {
            if (records.has('DPB')) { data = decodeProtectionRecord(records.get('DPB')!); }
            if (records.has('CMG')) {
                const state = decodeProtectionRecord(records.get('CMG')!);
                if (state.length !== 4) { throw new Error(INVALID); }
                locked ||= (state.readUInt32LE() & 7) !== 0;
            }
            if (records.has('GC')) {
                const visible = decodeProtectionRecord(records.get('GC')!);
                if (visible.length !== 1 || (visible[0] !== 0 && visible[0] !== 255)) { throw new Error(INVALID); }
                locked ||= visible[0] === 0;
            }
            if (data && !(data.length === 1 && data[0] === 0)) {
                if (data.length === 29 && data[0] === 255) {
                    if (data[28] !== 0) { throw new Error(INVALID); }
                } else if (data.length < 2 || data[data.length - 1] !== 0 || data.subarray(0, -1).includes(0)) {
                    throw new Error(INVALID);
                }
            }
        } catch { invalid = true; }
        this.passwordData = data;
        this.invalid = invalid;
        this.hasPassword = invalid || !!(data && !(data.length === 1 && data[0] === 0));
        // Also refuse host/editor locks without a usable password record.
        this.requiresPassword = this.hasPassword || locked;
        this.fingerprint = createHash('sha256').update(String(codePage)).update(String(invalid)).update([...records].sort().join('|')).digest('hex');
    }

    verify(password: string): boolean {
        const data = this.passwordData;
        if (this.invalid || !data || !this.hasPassword) { throw new Error(INVALID); }
        const encoded = encodeCodePage(password, this.codePage);
        // Refuse lossy encoding (e.g. emoji folding to '?' in an ANSI project).
        if (password.includes('\0') || decodeCodePage(encoded, this.codePage) !== password) { return false; }
        if (data.length === 29 && data[0] === 255) {
            const restored = Buffer.from(data.subarray(4, 28));
            for (let i = 0; i < 24; i++) {
                if (!(data[1 + Math.floor(i / 8)] & (128 >> (i % 8)))) { restored[i] = 0; }
            }
            const actual = createHash('sha1').update(encoded).update(restored.subarray(0, 4)).digest();
            return actual.equals(restored.subarray(4));
        }
        return encoded.equals(data.subarray(0, -1));
    }
}

// Only fingerprints live here. Passwords never leave the masked input flow,
// and are not saved to settings, logs, secrets, or the container.
const authorized = new Map<string, string>();

export function assertVbaProjectAccess(filePath: string | undefined, protection: VbaProjectProtection, revokeOnMismatch = true): void {
    if (!protection.requiresPassword) {
        if (filePath && revokeOnMismatch) { authorized.delete(projectIdentityKey(filePath)); }
        return;
    }
    if (filePath && authorized.get(projectIdentityKey(filePath)) === protection.fingerprint) { return; }
    if (filePath && revokeOnMismatch) { authorized.delete(projectIdentityKey(filePath)); }
    throw new VbaProjectLockedError();
}

export function authorizeVbaProject(filePath: string, protection: VbaProjectProtection, password: string): boolean {
    if (!protection.requiresPassword) { return true; }
    if (!protection.verify(password)) { return false; }
    authorized.set(projectIdentityKey(filePath), protection.fingerprint);
    return true;
}

export function clearVbaProjectAuthorizations(): void { authorized.clear(); }
