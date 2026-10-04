import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
const folder = (): XlideNode => ({ kind: 'shapes', shapeFolder: 'surface', surface: 'Data', label: 'Shapes', filePath: BOOK });
const listing = (name: string) => ({ surfaces: [{ surface: 'Data', shapes: [
    { name: 'Pair', kind: 'group', shapes: [{ name: 'Nested', kind: 'group', shapes: [{ name, kind: 'shape' }] }] },
] }] });
const create = (call: ReturnType<typeof vi.fn>) => new ShapeRows({ call } as unknown as ProjectEngine, () => {});

describe('retained shape groups after refresh', () => {
    it('reads current members before the parent folder redraws', async () => {
        const call = vi.fn().mockResolvedValue(listing('Old')), tree = create(call);
        const [group] = await tree.children(folder(), async () => []);
        const [nested] = await tree.children(group, async () => []);
        expect((await tree.children(nested, async () => [])).map(row => row.label)).toEqual(['Old']);
        call.mockResolvedValue(listing('New'));
        tree.refresh(BOOK);
        const [member] = await tree.children(nested, async () => []);
        expect(member.label).toBe('New');
        expect(member.shapePath).toEqual(['Pair', 'Nested', 'New']);
        expect(tree.parentOf(member)).toBe(nested);
        expect(tree.contextOf(member)).toMatchObject({ surface: 'Data', inGroup: true });
        expect(call).toHaveBeenCalledTimes(2);
    });

    it('does not resurrect members of a deleted group', async () => {
        const call = vi.fn().mockResolvedValue(listing('Old')), tree = create(call);
        const [group] = await tree.children(folder(), async () => []);
        call.mockResolvedValue({ surfaces: [{ surface: 'Data', shapes: [] }] });
        tree.refresh(BOOK);
        expect(await tree.children(group, async () => [])).toEqual([]);
    });

    it('keeps warm group expansion free of bridge reads and top-level scans', async () => {
        let scans = 0;
        const response = listing('Old');
        const call = vi.fn().mockResolvedValue(response), tree = create(call);
        const [group] = await tree.children(folder(), async () => []);
        Object.defineProperty(response.surfaces[0], 'shapes', { get: () => { scans++; return []; } });
        for (let i = 0; i < 200; i++) {
            expect((await tree.children(group, async () => [])).map(row => row.label)).toEqual(['Nested']);
        }
        expect(scans).toBe(0);
        expect(call).toHaveBeenCalledTimes(1);
    });

    it('follows the module code name when a grouped worksheet is renamed', async () => {
        const original = listing('Old');
        Object.assign(original.surfaces[0], { codeName: 'Sheet1' });
        const call = vi.fn().mockResolvedValue(original), tree = create(call);
        const moduleFolder: XlideNode = { ...folder(), shapeFolder: 'module', moduleName: 'Sheet1' };
        const [group] = await tree.children(moduleFolder, async () => []);
        const renamed = listing('New');
        Object.assign(renamed.surfaces[0], { codeName: 'Sheet1', surface: 'Renamed' });
        call.mockResolvedValue(renamed);
        tree.refresh(BOOK);
        const [nested] = await tree.children(group, async () => []);
        expect(nested.surface).toBe('Renamed');
        expect(tree.contextOf(group)).toMatchObject({ surface: 'Renamed' });
        expect((await tree.children(nested, async () => [])).map(row => row.label)).toEqual(['New']);
    });

    it('does not replace a group snapshot after disposal during its read', async () => {
        let resolve!: (value: unknown) => void;
        const pendingListing = new Promise(yes => { resolve = yes; });
        const call = vi.fn().mockResolvedValue(listing('Old')), tree = create(call);
        const [group] = await tree.children(folder(), async () => []);
        const original = group.shape;
        call.mockReturnValue(pendingListing);
        tree.refresh(BOOK);
        const pending = tree.children(group, async () => []);
        tree.dispose();
        resolve(listing('New'));
        expect(await pending).toEqual([]);
        expect(group.shape).toBe(original);
    });

    it.each(['success', 'failure'] as const)('joins a current group read after an overtaken %s', async outcome => {
        let resolve!: (value: unknown) => void, reject!: (error: Error) => void;
        const old = new Promise((yes, no) => { resolve = yes; reject = no; });
        const call = vi.fn().mockResolvedValue(listing('Old')), tree = create(call);
        const [group] = await tree.children(folder(), async () => []);
        const [nested] = await tree.children(group, async () => []);
        call.mockReturnValueOnce(old).mockResolvedValue(listing('New'));
        tree.refresh(BOOK);
        const pending = tree.children(nested, async () => []);
        tree.refresh(BOOK);
        expect((await tree.children(nested, async () => [])).map(row => row.label)).toEqual(['New']);
        if (outcome === 'success') { resolve(listing('Obsolete')); }
        else { reject(new Error('Obsolete read')); }
        expect((await pending).map(row => row.label)).toEqual(['New']);
        expect(call).toHaveBeenCalledTimes(3);
    });
});
