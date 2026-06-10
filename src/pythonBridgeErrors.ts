import { errorMessage } from './util/errors';

/**
 * Recognizes the bridge rejection for backends (and test fakes) that do not
 * implement the batch readModules RPC, so callers can fall back to
 * listModules plus per-module readModule calls.
 */
export function isReadModulesUnavailable(err: unknown): boolean {
    const message = errorMessage(err);
    return /Method not found:\s*readModules/i.test(message) ||
        /Unexpected bridge call readModules/i.test(message);
}
