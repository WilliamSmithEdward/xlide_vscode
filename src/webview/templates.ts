import { readExtensionTextAsset } from '../extensionAssets';

/**
 * Fills a webview template/asset (assets/webview/*) loaded from disk at
 * runtime. Templates mark dynamic regions with {{placeholder}} tokens; every
 * token must have a substitution and every file is expected to end with a
 * single trailing newline (trimmed here so embedding a .css/.js body between
 * template lines reproduces the original inline layout byte for byte).
 *
 * Substitution values are inserted verbatim and never re-scanned, so model
 * JSON or user content containing '{{' cannot corrupt the output.
 */
export function renderWebviewTemplate(
    assetRelativePath: string,
    substitutions: Readonly<Record<string, string>>,
): string {
    const template = readExtensionTextAsset(assetRelativePath).replace(/\n$/, '');
    return template.replace(/\{\{([A-Za-z0-9]+)\}\}/g, (token, key: string) => {
        const value = substitutions[key];
        if (value === undefined) {
            throw new Error(`Webview template ${assetRelativePath}: no substitution for ${token}`);
        }
        return value;
    });
}
