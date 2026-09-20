// Getting a container's bytes to the engine, and its writes back out.
//
// The container engine is synchronous: it reads a whole file, works in
// memory, and writes it back (see src/vba/hostPlatform.ts for why that does
// not change). On a desktop it simply calls node:fs and there is nothing to
// arrange. In a browser there is no synchronous file access at all, so the
// bytes have to be in hand before the engine runs.
//
// ProjectEngine.call() is the only place that reaches the engine, and it is
// already async, so the whole arrangement fits around that one call: load
// what the call will read, run it, send back whatever it wrote. A failed call
// discards its writes rather than flushing a half-finished container.

import { enginePriming as leaf } from './enginePrimingNode';

export interface EnginePriming {
    /** Loads what the engine is about to read. */
    prime(filePaths: readonly string[]): Promise<void>;

    /** Sends everything the engine wrote back to the workspace. */
    flush(): Promise<void>;

    /**
     * Throws away everything the engine wrote. Used when the call failed:
     * the engine rewrites a container whole, so a partial write is not a
     * partial file, it is a file that was never finished.
     */
    discard(): void;
}

export const enginePriming: EnginePriming = leaf;
