import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';

interface PackageToolContribution {
    name: string;
    toolReferenceName: string;
    modelDescription?: string;
    inputSchema?: {
        required?: string[];
        properties?: Record<string, unknown>;
    };
}

function languageModelTools(): PackageToolContribution[] {
    const manifest = JSON.parse(readFileSync('package.json', 'utf8')) as {
        contributes?: { languageModelTools?: PackageToolContribution[] };
    };
    return manifest.contributes?.languageModelTools ?? [];
}

describe('XLIDE agent tool manifest', () => {
    it('exposes project analysis to AI agents', () => {
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_analyzeProject');

        expect(tool).toEqual(expect.objectContaining({
            toolReferenceName: 'xlideAnalyzeProject',
        }));
        expect(tool?.inputSchema?.required).toContain('filePath');
    });

    it('tells agents that createProject never overwrites an existing file', () => {
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_createProject');

        expect(tool?.modelDescription).toContain('never overwrites');
        expect(tool?.modelDescription).not.toContain('Overwrites the file');
        expect(tool?.inputSchema?.required).toContain('filePath');
    });

    it('exposes VBA test execution to AI agents', () => {
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_runVbaTests');

        expect(tool).toEqual(expect.objectContaining({
            toolReferenceName: 'xlideRunVbaTests',
        }));
        expect(tool?.modelDescription).toContain('artifacts');
        expect(tool?.modelDescription).toContain('status_for_ci.json');
        expect(tool?.inputSchema?.required).toContain('filePath');
        expect(tool?.inputSchema?.properties).toEqual(expect.objectContaining({
            includeTags: expect.any(Object),
            excludeTags: expect.any(Object),
            failFast: expect.any(Object),
            includeHostEvents: expect.any(Object),
        }));
    });
});
