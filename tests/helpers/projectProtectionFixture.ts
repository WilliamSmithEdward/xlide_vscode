// Deterministic test-only MS-OVBA records. This fixture password is public.
import { createHash } from 'crypto';
import * as fs from 'fs';
import { encodeCodePage } from '../../src/vba/codePages';
import { openMacroContainer } from '../../src/vba/macroContainer';
import { resetProjectCacheForTests } from '../../src/vba/projectService';
import { AccessPageStore } from '../../src/vba/access/accessPageStore';
import { AccessTable } from '../../src/vba/access/accessTableWriter';
import { readAccessCatalog, readAccessStorage, type AccessStorageEntry } from '../../src/vba/access/accessStorage';
export const FIXTURE_PASSWORD = 'Test66';

export function record(data: Buffer, seed = 6): string {
    const bytes = [seed, seed ^ 2, seed ^ 0xac];
    let plain = 0xac, last = bytes[2], older = bytes[1];
    const length = Buffer.alloc(4); length.writeUInt32LE(data.length);
    const input = Buffer.concat([Buffer.alloc((seed & 6) >> 1, 0x71), length, data]);
    for (const value of input) {
        const encrypted = value ^ ((older + plain) & 255);
        bytes.push(encrypted); older = last; last = encrypted; plain = value;
    }
    return Buffer.from(bytes).toString('hex').toUpperCase();
}
export function hashed(password: string, cp = 1252): Buffer {
    const key = Buffer.from([0, 0x34, 0, 0x92]);
    const digest = createHash('sha1').update(encodeCodePage(password, cp)).update(key).digest();
    const input = Buffer.concat([key, digest]);
    const result = Buffer.alloc(29); result[0] = 255;
    for (let i = 0; i < 24; i++) {
        if (input[i]) { result[1 + Math.floor(i / 8)] |= 128 >> (i % 8); }
        result[4 + i] = input[i] || 1;
    }
    return result;
}
export function raw(password: string, cp = 1252, legacy = false): Buffer {
    const data = legacy ? Buffer.concat([encodeCodePage(password, cp), Buffer.from([0])]) : hashed(password, cp);
    return Buffer.from(`CMG="${record(Buffer.from([4, 0, 0, 0]))}"\r\nDPB="${record(data)}"\r\nGC="${record(Buffer.from([0]))}"\r\n`);
}
export function protectFixture(target: string, password: string): void {
    const bytes = fs.readFileSync(target);
    const container = openMacroContainer(bytes);
    const cfb = container.vbaCfb();
    const previous = cfb.getStream('PROJECT').toString('latin1');
    const header = previous.replace(/^(?:CMG|DPB|GC)=.*\r?\n/gm, '')
        .replace(/^ID=.*$/m, 'ID="{00000000-0000-0000-0000-000000000000}"');
    // PROJECT's protection belongs before the first section header. Appending
    // it beneath [Workspace] produces records the native reader sees but Office
    // ignores, so the fixture would not actually be locked in Excel.
    const section = header.search(/^\[/m);
    const project = Buffer.from(section < 0 ? header + raw(password).toString('latin1')
        : header.slice(0, section) + raw(password).toString('latin1') + header.slice(section), 'latin1');
    if (container.kind === 'access') {
        // Access's production writer intentionally preserves protection. For
        // test setup, patch the actual named storage row, not its synthetic CFB.
        const visit = (entries: AccessStorageEntry[]): AccessStorageEntry[] => entries.flatMap(entry => [entry, ...visit(entry.children)]);
        const stream = visit(readAccessStorage(bytes) ?? []).find(entry => entry.name === 'PROJECT');
        const catalog = readAccessCatalog(bytes).find(entry => entry.name === 'MSysAccessStorage');
        if (!stream || !catalog) { throw new Error('Access fixture project storage missing'); }
        const store = new AccessPageStore(bytes);
        const table = new AccessTable(store, catalog.definitionPage);
        table.updateNamedRow({ page: stream.page, slot: stream.slot }, new Map([['Lv', project]]));
        fs.writeFileSync(target, store.toBuffer());
    } else {
        cfb.writeStream('PROJECT', project);
        fs.writeFileSync(target, container.toFileBytes(cfb));
    }
    resetProjectCacheForTests();
}
