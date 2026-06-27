import { describe, it, expect } from 'vitest';
import { createKeyedAsyncLock } from '../src/util/keyedAsyncLock';

const tick = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

describe('createKeyedAsyncLock', () => {
    it('serializes actions sharing a key in call order', async () => {
        const lock = createKeyedAsyncLock();
        const order: string[] = [];
        const a = lock('k', async () => {
            order.push('a-start');
            await tick();
            order.push('a-end');
        });
        const b = lock('k', async () => {
            order.push('b-start');
            await tick();
            order.push('b-end');
        });
        await Promise.all([a, b]);
        // b must not start until a has fully finished.
        expect(order).toEqual(['a-start', 'a-end', 'b-start', 'b-end']);
    });

    it('runs different keys concurrently', async () => {
        const lock = createKeyedAsyncLock();
        const order: string[] = [];
        const a = lock('k1', async () => {
            order.push('a-start');
            await tick();
            order.push('a-end');
        });
        const b = lock('k2', async () => {
            order.push('b-start');
            await tick();
            order.push('b-end');
        });
        await Promise.all([a, b]);
        // Both start before either ends (independent queues).
        expect(order.slice(0, 2).sort()).toEqual(['a-start', 'b-start']);
    });

    it('does not wedge the queue when an action throws', async () => {
        const lock = createKeyedAsyncLock();
        await expect(lock('k', async () => { throw new Error('boom'); })).rejects.toThrow('boom');
        await expect(lock('k', async () => 'ok')).resolves.toBe('ok');
    });

    it('returns the action result', async () => {
        const lock = createKeyedAsyncLock();
        await expect(lock('k', async () => 42)).resolves.toBe(42);
    });
});
