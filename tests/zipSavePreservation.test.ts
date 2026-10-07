import * as fs from 'fs';
import * as path from 'path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { containerCodec, setContainerCodec } from '../src/vba/containerCodec';
import { openMacroContainer } from '../src/vba/macroContainer';
import { ZipArchive } from '../src/vba/zip';

const codec = containerCodec();
afterEach(() => setContainerCodec(codec));

// pyOpenVBA #56 avoids inflating and recompressing unchanged Office parts.
// XLIDE already uses that path; guard its work count as well as its bytes.
describe('Office saves retain untouched compressed parts', () => {
    it.each(['SheetsFixture.xlsm', 'WordFixture.docm', 'PowerPointFixture.pptm'])(
        '%s compresses only the VBA project, including on repeated saves', (file) => {
            const seed = ZipArchive.read(fs.readFileSync(path.join(__dirname, 'fixtures', 'binaries', file)));
            const attachment = Buffer.from('<row><cell>untouched content</cell></row>'.repeat(100000));
            seed.write('attachment.xml', attachment);
            const original = seed.toBytes();
            const before = storedParts(original);
            const container = openMacroContainer(original);
            const cfb = container.vbaCfb();
            const projectPath = seed.names().find((name) => name.endsWith('/vbaProject.bin'))!;
            expect(projectPath).toBeDefined();
            const project = cfb.getStream('PROJECT');
            cfb.writeStream('PROJECT', Buffer.concat([project, Buffer.from('\r\n; save regression\r\n')]));
            const expectedProject = cfb.toBytes();

            const inflateRaw = vi.fn(codec.inflateRaw);
            const compressForZip = vi.fn(codec.compressForZip);
            setContainerCodec({ ...codec, inflateRaw, compressForZip });
            const first = container.toFileBytes(cfb);
            const second = container.toFileBytes(cfb);
            expect(inflateRaw).not.toHaveBeenCalled();
            expect(compressForZip).toHaveBeenCalledTimes(2);
            for (const [data] of compressForZip.mock.calls) {
                expect(data.equals(expectedProject)).toBe(true);
            }
            setContainerCodec(codec);

            for (const saved of [first, second]) {
                const after = storedParts(saved);
                expect([...after.keys()]).toEqual([...before.keys()]);
                for (const [name, compressed] of before) {
                    if (name !== projectPath) {
                        expect(after.get(name)!.equals(compressed), name).toBe(true);
                    }
                }
                const reopened = ZipArchive.read(saved);
                expect(reopened.read(projectPath).equals(expectedProject)).toBe(true);
                expect(reopened.read('attachment.xml').equals(attachment)).toBe(true);
            }
        },
    );
});

/** Read compressed payloads directly, without the production ZIP reader. */
function storedParts(bytes: Buffer): Map<string, Buffer> {
    const end = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    let at = bytes.readUInt32LE(end + 16);
    const parts = new Map<string, Buffer>();
    for (let i = 0; i < bytes.readUInt16LE(end + 10); i++) {
        expect(bytes.readUInt32LE(at)).toBe(0x02014b50);
        const size = bytes.readUInt32LE(at + 20);
        const nameLength = bytes.readUInt16LE(at + 28);
        const name = bytes.subarray(at + 46, at + 46 + nameLength).toString('utf8');
        const local = bytes.readUInt32LE(at + 42);
        const start = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28);
        parts.set(name, bytes.subarray(start, start + size));
        at += 46 + nameLength + bytes.readUInt16LE(at + 30) + bytes.readUInt16LE(at + 32);
    }
    return parts;
}
