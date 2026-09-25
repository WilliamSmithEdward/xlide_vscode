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

    it('steers agents to write VBA through the tool, which is what gets the user s review', () => {
        // An edit an agent makes to an open xlide-vba:// document is saved into
        // the file like any other, and XLIDE cannot tell it from the user's own
        // typing, so it gets no diff and no tree badge. The description used to
        // offer that as an equal path.
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_writeModule');

        expect(tool?.modelDescription).toContain('Make every change with this tool, even to a module the user has open');
        expect(tool?.modelDescription).toContain('no before/after diff and no tree badge');
        expect(tool?.modelDescription).not.toContain('persist only through this tool or through the XLIDE virtual file system');
    });

    it('reads parts of a module in one call, and requires the read s token to edit them', () => {
        const read = languageModelTools().find((entry) => entry.name === 'xlide_readModule');
        expect(Object.keys(read?.inputSchema?.properties ?? {})).toEqual(expect.arrayContaining(['ranges', 'procedures']));

        const edit = languageModelTools().find((entry) => entry.name === 'xlide_editModule');
        expect(edit).toEqual(expect.objectContaining({ toolReferenceName: 'xlideEditModule' }));
        // Line numbers mean one read; without its token an edit lands on
        // whatever the module holds now.
        expect(edit?.inputSchema?.required).toEqual(['filePath', 'moduleName', 'expectedContentToken', 'edits']);
        expect(edit?.modelDescription).toContain('as xlide_readModule last showed it');
        expect(edit?.modelDescription).toContain('Keep/Revert');
    });

    it('imports a folder of module files on request, and says that is not the way to edit VBA', () => {
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_importModules');

        expect(tool).toEqual(expect.objectContaining({ toolReferenceName: 'xlideImportModules' }));
        expect(tool?.inputSchema?.required).toEqual(['filePath']);
        expect((tool?.inputSchema?.properties?.importMode as { enum?: string[] })?.enum)
            .toEqual(['updateOnly', 'trueUpStandardClass']);
        expect(tool?.modelDescription).toContain('not the way to edit VBA');
        expect(tool?.modelDescription).toContain('confirmation');
    });

    it('offers the libraries a project can be given a reference to, and only those', () => {
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_addReference');

        expect(tool).toEqual(expect.objectContaining({
            toolReferenceName: 'xlideAddReference',
        }));
        expect(tool?.inputSchema?.required).toEqual(['filePath', 'library']);
        expect((tool?.inputSchema?.properties?.library as { enum?: string[] })?.enum)
            .toEqual(['excel', 'word', 'powerpoint', 'access']);
        // Late binding needs no reference, and an agent that adds one for code
        // that would run without it has changed the project for nothing.
        expect(tool?.modelDescription).toContain('Late binding');
    });

    it('tells an agent that removing a reference breaks the code that names it', () => {
        const tool = languageModelTools().find((entry) => entry.name === 'xlide_removeReference');

        expect(tool).toEqual(expect.objectContaining({
            toolReferenceName: 'xlideRemoveReference',
        }));
        expect(tool?.inputSchema?.required).toEqual(['filePath', 'library']);
        // Not an enum like the add tool's: a project's references are not only
        // the four applications XLIDE can add one for.
        expect((tool?.inputSchema?.properties?.library as { enum?: string[] })?.enum).toBeUndefined();
        expect(tool?.modelDescription).toContain('stops compiling');
    });

    it('names every tool in the repository agent instructions', () => {
        // The instructions still said 18 tools after the 19th and 20th
        // shipped, and neither had a row in their tables.
        const instructions = readFileSync('.github/copilot-instructions.md', 'utf8');
        const missing = languageModelTools()
            .map((tool) => tool.name)
            .filter((name) => !instructions.includes(`\`${name}\``));

        expect(missing).toEqual([]);
        expect(instructions).not.toMatch(/\b\d+ tools\b/);
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
