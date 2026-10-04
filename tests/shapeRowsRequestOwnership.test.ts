import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
const project: XlideNode = { kind: 'project', label: 'Book', filePath: BOOK };
const sheets = [{ name: 'Data', kind: 'worksheet' as const }];
function deferred() {
    let resolve!: (value: unknown) => void, reject!: (error: unknown) => void;
    const promise = new Promise<unknown>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}
function create(call = vi.fn().mockResolvedValue({ sheets, surfaces: [] })) {
    const appendLine = vi.fn(), fire = vi.fn();
    const tree = new ShapeRows({ call } as unknown as ProjectEngine, fire,
        { appendLine } as unknown as ConstructorParameters<typeof ShapeRows>[2]);
    return { tree, call, appendLine, fire };
}
const flush = async () => { for (let i = 0; i < 8; i++) { await Promise.resolve(); } };

describe('shape row request ownership', () => {
    it.each(['refresh', 'clear'] as const)('ignores obsolete sheet failure output after %s', async action => {
        const old = deferred();
        const { tree, call, appendLine } = create(vi.fn().mockReturnValueOnce(old.promise).mockResolvedValue({ sheets }));
        const pending = tree.catalog(BOOK);
        if (action === 'refresh') { tree.refresh(BOOK); } else { tree.clear(); }
        expect(await tree.catalog(BOOK)).toBe(sheets);
        old.reject(new Error('obsolete sheets'));
        expect(await pending).toBeUndefined();
        expect(appendLine).toHaveBeenCalledTimes(0);
        expect(await tree.catalog(BOOK)).toBe(sheets);
        expect(call).toHaveBeenCalledTimes(2);
    });

    it.each(['refresh', 'clear'] as const)('ignores obsolete shape failure output after %s', async action => {
        const old = deferred(), surfaces = [{ surface: 'Data', shapes: [] }];
        const { tree, call, appendLine } = create(vi.fn().mockReturnValueOnce(old.promise).mockResolvedValue({ surfaces }));
        const pending = tree.surfaces(BOOK).catch(error => error);
        if (action === 'refresh') { tree.refresh(BOOK); } else { tree.clear(); }
        expect(await tree.surfaces(BOOK)).toBe(surfaces);
        const error = new Error('obsolete shapes'); old.reject(error);
        expect(await pending).toBe(error);
        expect(appendLine).toHaveBeenCalledTimes(0);
        expect(await tree.surfaces(BOOK)).toBe(surfaces);
        expect(call).toHaveBeenCalledTimes(2);
    });

    it('does not let an obsolete failure replace the current sheet-warning deduplication reason', async () => {
        const old = deferred();
        const { tree, appendLine } = create(vi.fn()
            .mockRejectedValueOnce(new Error('reason A'))
            .mockReturnValueOnce(old.promise)
            .mockRejectedValue(new Error('reason B')));
        await tree.catalog(BOOK);
        tree.refresh(BOOK); const pending = tree.catalog(BOOK);
        tree.refresh(BOOK); await tree.catalog(BOOK);
        old.reject(new Error('reason A')); await pending;
        tree.refresh(BOOK); await tree.catalog(BOOK);
        expect(appendLine).toHaveBeenCalledTimes(2);
        expect(appendLine.mock.calls.map(args => args[0])).toEqual([
            expect.stringContaining('reason A'), expect.stringContaining('reason B'),
        ]);
    });

    it('does not redraw a discarded Sheets folder when its refresh read finishes after clear', async () => {
        const old = deferred();
        const { tree, fire } = create(vi.fn().mockResolvedValueOnce({ sheets }).mockResolvedValueOnce({ surfaces: [] }).mockReturnValueOnce(old.promise));
        const { folders: [folder] } = await tree.projectRows(project, []);
        await tree.children(folder, async () => []);
        tree.refresh(BOOK); tree.clear();
        old.resolve({ sheets: [...sheets, { name: 'Added', kind: 'worksheet' }] }); await flush();
        expect(fire).toHaveBeenCalledTimes(0);
    });

    it('does not redraw from a sheet refresh overtaken by a newer refresh', async () => {
        const old = deferred();
        const { tree, fire, call } = create(vi.fn().mockResolvedValueOnce({ sheets }).mockResolvedValueOnce({ surfaces: [] }).mockReturnValueOnce(old.promise).mockResolvedValue({ sheets }));
        const { folders: [folder] } = await tree.projectRows(project, []);
        await tree.children(folder, async () => []);
        tree.refresh(BOOK); tree.refresh(BOOK); await flush();
        old.resolve({ sheets: [...sheets, { name: 'Obsolete', kind: 'worksheet' }] }); await flush();
        expect(fire).toHaveBeenCalledTimes(0);
        expect(await tree.catalog(BOOK)).toBe(sheets);
        expect(call).toHaveBeenCalledTimes(4);
    });

    it('still redraws a current Sheets folder when its sheet catalog changes', async () => {
        const { tree, fire } = create(vi.fn().mockResolvedValueOnce({ sheets }).mockResolvedValueOnce({ surfaces: [] }).mockResolvedValue({ sheets: [...sheets, { name: 'Added', kind: 'worksheet' }] }));
        const { folders: [folder] } = await tree.projectRows(project, []);
        await tree.children(folder, async () => []);
        tree.refresh(BOOK); await flush();
        expect(fire.mock.calls).toEqual([[folder]]);
    });

    it('skips output work for a burst of 100 superseded sheet and shape requests', async () => {
        const obsolete = Array.from({ length: 200 }, () => deferred());
        let index = 0;
        const { tree, call, appendLine } = create(vi.fn().mockImplementation(() =>
            index < obsolete.length ? obsolete[index++].promise : Promise.resolve({ sheets, surfaces: [] })));
        const pending: Promise<unknown>[] = [];
        for (let i = 0; i < 100; i++) {
            tree.refresh(BOOK);
            pending.push(tree.catalog(BOOK), tree.surfaces(BOOK).catch(error => error));
        }
        tree.refresh(BOOK);
        await Promise.all([tree.catalog(BOOK), tree.surfaces(BOOK)]);
        obsolete.forEach((request, i) => request.reject(new Error('obsolete ' + i)));
        await Promise.all(pending);
        expect(appendLine).toHaveBeenCalledTimes(0);
        expect(call).toHaveBeenCalledTimes(202);
    });

    it('coalesces and negative-caches a current shape failure, then retries after refresh', async () => {
        const error = new Error('current failure');
        const { tree, call, appendLine } = create(vi.fn().mockRejectedValue(error));
        const results = await Promise.all(Array.from({ length: 20 }, () => tree.surfaces(BOOK).catch(value => value)));
        expect(results.every(value => value === error)).toBe(true);
        await tree.surfaces(BOOK).catch(() => {});
        expect(call).toHaveBeenCalledTimes(1);
        expect(appendLine).toHaveBeenCalledTimes(1);
        tree.refresh(BOOK); await tree.surfaces(BOOK).catch(() => {});
        expect(call).toHaveBeenCalledTimes(2);
        expect(appendLine).toHaveBeenCalledTimes(2);
    });
});
