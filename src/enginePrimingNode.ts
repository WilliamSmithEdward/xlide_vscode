// On a desktop there is nothing to arrange: the engine reads and writes the
// real filesystem itself, synchronously, exactly as it always has. Every
// member here is deliberately empty, and this file exists so that
// ProjectEngine.call() has one shape on both platforms.
//
// The browser build never reaches this module - webBuild.js aliases it to
// enginePrimingWeb.ts.

import type { EnginePriming } from './enginePriming';

export const enginePriming: EnginePriming = {
    async prime(_filePaths: readonly string[]): Promise<void> {
        /* node:fs reads the file when the engine asks */
    },

    async flush(): Promise<void> {
        /* atomicWrite already landed the bytes */
    },

    discard(): void {
        /* nothing was buffered */
    },
};
