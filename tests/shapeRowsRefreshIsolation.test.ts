import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';

const TARGET = '/work/Target.pptm';
const row = (filePath: string, surface: string): XlideNode => ({ kind: 'surface', label: surface, filePath, surface });
function fixture(call = vi.fn(async () => ({ surfaces: [] }))) {
    const fire = vi.fn();
    return { tree: new ShapeRows({ call } as unknown as ProjectEngine, fire), call, fire };
}

describe('shape refresh project isolation', () => {
    it('does no work on a thousand opened surfaces in another project', async () => {
        const { tree, call, fire } = fixture();
        let unrelatedPathReads = 0;
        for (let i = 0; i < 1000; i++) {
            const other = row('/work/Other.pptm', `Slide ${i}`);
            Object.defineProperty(other, 'filePath', { get: () => { unrelatedPathReads++; return '/work/Other.pptm'; } });
            await tree.children(other, async () => []);
        }
        const target = row(TARGET, 'Slide 1');
        await tree.children(target, async () => []);
        unrelatedPathReads = 0;
        const readsBefore = call.mock.calls.length;
        for (let i = 0; i < 10; i++) { tree.refresh(TARGET); }
        expect(unrelatedPathReads).toBe(0);
        expect(fire.mock.calls.map(([node]) => node)).toEqual(Array(10).fill(target));
        expect(call.mock.calls.length).toBe(readsBefore);
    });

    it('notifies each opened row once using normalized project identity', async () => {
        const { tree, fire } = fixture();
        const first = row(TARGET, 'Slide 1'), second = row('/work/./Target.pptm', 'Slide 2');
        for (let i = 0; i < 3; i++) { await tree.children(first, async () => []); }
        await tree.children(second, async () => []);
        tree.refresh('/work/unused/../Target.pptm');
        expect(fire.mock.calls.map(([node]) => node)).toEqual([first, second]);
    });

    it('forgets opened rows on a full clear and after disposal', async () => {
        const { tree, call, fire } = fixture();
        await tree.children(row(TARGET, 'Slide 1'), async () => []);
        tree.clear();
        tree.refresh(TARGET);
        expect(fire).not.toHaveBeenCalled();
        await tree.children(row(TARGET, 'Slide 2'), async () => []);
        tree.dispose();
        const before = call.mock.calls.length;
        tree.refresh(TARGET);
        expect(fire).not.toHaveBeenCalled();
        expect(call.mock.calls.length).toBe(before);
    });

    it('refreshes rows whose previous shapes read failed so expansion can retry', async () => {
        const call = vi.fn().mockRejectedValueOnce(new Error('Busy')).mockResolvedValue({ surfaces: [] });
        const { tree, fire } = fixture(call);
        const target = row(TARGET, 'Slide 1');
        expect((await tree.children(target, async () => []))[0].label).toBe('Shapes could not be read');
        tree.refresh(TARGET);
        expect(fire).toHaveBeenCalledWith(target);
        expect((await tree.children(target, async () => []))[0].label).toBe('No shapes');
        expect(call).toHaveBeenCalledTimes(2);
    });
});
