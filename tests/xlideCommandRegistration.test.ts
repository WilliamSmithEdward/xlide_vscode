import { describe, it, expect, vi, beforeEach } from 'vitest';

const hoisted = vi.hoisted(() => ({ registerCommand: vi.fn() }));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    commands: { registerCommand: hoisted.registerCommand },
}));

import { registerXlideCommand } from '../src/xlideCommandRegistration';

describe('registerXlideCommand', () => {
    beforeEach(() => {
        hoisted.registerCommand.mockReset();
    });

    it('returns the disposable from vscode.commands.registerCommand on success', () => {
        const disposable = { dispose: vi.fn() };
        hoisted.registerCommand.mockReturnValue(disposable);

        const result = registerXlideCommand('xlide.test', () => undefined);

        expect(result).toBe(disposable);
        expect(hoisted.registerCommand).toHaveBeenCalledWith('xlide.test', expect.any(Function));
    });

    it('survives a duplicate-id registration instead of throwing (activation-resilience fix)', () => {
        // A stale registration left by a prior partial activation makes
        // registerCommand throw "already exists". This must NOT propagate, or it
        // would abort registration of the remaining commands during activation.
        const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
        hoisted.registerCommand.mockImplementation(() => {
            throw new Error("command 'xlide.test' already exists");
        });

        let result: { dispose: () => void } | undefined;
        expect(() => { result = registerXlideCommand('xlide.test', () => undefined); }).not.toThrow();
        expect(result).toBeDefined();
        // The returned no-op disposable must be safely disposable.
        expect(() => result?.dispose()).not.toThrow();
        expect(warn).toHaveBeenCalled();

        warn.mockRestore();
    });
});
