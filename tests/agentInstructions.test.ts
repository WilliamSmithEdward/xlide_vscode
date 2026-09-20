import { readFileSync } from 'fs';
import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    window: { activeTextEditor: undefined, visibleTextEditors: [] },
    workspace: { workspaceFolders: [] },
}));

import * as vscode from 'vscode';
import { AGENT_INSTRUCTIONS, AGENT_INSTRUCTIONS_STEPS } from '../src/agentInstructions';
import { renderXlideSidebarHtml, XlideSidebarProvider } from '../src/xlideSidebar';
import { buildXlideSidebarModel } from '../src/xlideSidebarModel';

interface Manifest {
    contributes: {
        languageModelTools: Array<{ name: string; toolReferenceName: string }>;
        commands: Array<{ command: string; title: string; category?: string }>;
    };
}

const manifest = JSON.parse(readFileSync('package.json', 'utf8')) as Manifest;
const tools = manifest.contributes.languageModelTools;

function decodeHtml(text: string): string {
    return text.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
}

describe('the agent instructions text', () => {
    it('names every tool XLIDE registers', () => {
        // A tool the text leaves out is one no agent reading it will reach for.
        const missing = tools.map((tool) => tool.name).filter((name) => !new RegExp(`\\b${name}\\b`).test(AGENT_INSTRUCTIONS));

        expect(missing).toEqual([]);
    });

    it('names no tool, chat reference or command that does not exist', () => {
        // An agent told to call a tool that is not there loses a turn finding out.
        // The MCP server section is cut out first: its tools are its own
        // namespace, named in snake_case, and are checked separately.
        const extensionTools = AGENT_INSTRUCTIONS.slice(0, AGENT_INSTRUCTIONS.indexOf('## The XLIDE MCP server'));
        const named = [...extensionTools.matchAll(/(?<![.\w])xlide_[A-Za-z]+\b/g)].map((m) => m[0]);
        const referenced = [...AGENT_INSTRUCTIONS.matchAll(/#(xlide[A-Za-z]+)/g)].map((m) => m[1]);
        const commands = [...AGENT_INSTRUCTIONS.matchAll(/"XLIDE: ([^"]+)"/g)].map((m) => m[1]);

        expect(named.length).toBeGreaterThan(0);
        expect(named.filter((name) => !tools.some((tool) => tool.name === name))).toEqual([]);
        expect(referenced.filter((name) => !tools.some((tool) => tool.toolReferenceName === name))).toEqual([]);
        expect(commands.length).toBeGreaterThan(0);
        expect(commands.filter((title) => !manifest.contributes.commands.some((c) => c.category === 'XLIDE' && c.title === title))).toEqual([]);
    });

    it('names the buttons the person has to press as XLIDE labels them', () => {
        const analysisView = readFileSync('assets/webview/projectAnalysis.html', 'utf8');
        const sidebar = JSON.stringify(buildXlideSidebarModel({}));

        for (const label of ['Copy JSON', 'Export JSON']) {
            expect(AGENT_INSTRUCTIONS).toContain(label);
            expect(analysisView).toContain(`>${label}</button>`);
        }
        for (const label of ['Unit Tests', 'Analyze Project']) {
            expect(AGENT_INSTRUCTIONS).toContain(label);
            expect(sidebar).toContain(`"label":"${label}"`);
        }
        for (const title of [
            'Add Standard Module', 'Add Class Module', 'Rename Module', 'Delete Module', 'Open Form Markup',
            'Export All Modules to Folder', 'Import Modules from Folder',
        ]) {
            expect(AGENT_INSTRUCTIONS).toContain(title);
            expect(manifest.contributes.commands.some((command) => command.title === title), title).toBe(true);
        }
    });

    it('recommends the author\'s MCP server to an agent without the tools, and holds it to asking first', () => {
        const server = AGENT_INSTRUCTIONS.slice(AGENT_INSTRUCTIONS.indexOf('## The XLIDE MCP server'));

        expect(AGENT_INSTRUCTIONS).toContain('the author of XLIDE recommends xlide-mcp');
        expect(AGENT_INSTRUCTIONS).toContain('Install or add nothing without asking');
        expect(server).toContain('Add nothing before the user agrees');
        expect(server).toContain('--root is the security boundary');
        expect(AGENT_INSTRUCTIONS_STEPS.join(' ')).toContain('XLIDE\'s author recommends his xlide-mcp MCP server');
    });

    it('leads with the uvx route, which installs nothing, and keeps pip as the alternative', () => {
        // uv brings its own Python, so this is the one configuration that
        // works on a machine with neither it nor the package. An agent told
        // only about pip asks the user to install something they need not.
        const server = AGENT_INSTRUCTIONS.slice(AGENT_INSTRUCTIONS.indexOf('## The XLIDE MCP server'));

        expect(server).toContain('"command": "uvx"');
        expect(server).toContain('"--from", "xlide-mcp[live]"');
        expect(server).toContain('pip install xlide-mcp');
        expect(server.indexOf('"command": "uvx"')).toBeLessThan(server.indexOf('pip install xlide-mcp'));
        expect(AGENT_INSTRUCTIONS_STEPS.join(' ')).toContain('Nothing has to be installed');
    });

    it('names no library the author has replaced with the MCP server', () => {
        // The three Python libraries were the recommended route before
        // xlide-mcp; a leftover mention sends an agent to install the wrong
        // thing.
        for (const gone of ['pyOpenVBA', 'pyopenvba', 'pyvbaanalysis', 'pyvbaharness']) {
            expect(AGENT_INSTRUCTIONS, gone).not.toContain(gone);
            expect(AGENT_INSTRUCTIONS_STEPS.join(' '), gone).not.toContain(gone);
        }
    });

    it("names the MCP server's tools in its own snake_case namespace", () => {
        // They are a different namespace from the extension's camelCase
        // tools, and the test above deliberately does not check them against
        // package.json. Here they only have to be self-consistent.
        const server = AGENT_INSTRUCTIONS.slice(AGENT_INSTRUCTIONS.indexOf('## The XLIDE MCP server'));
        const named = [...server.matchAll(/\bxlide_[a-z_]+\b/g)].map((m) => m[0]);

        expect(named).toContain('xlide_read_module');
        expect(named).toContain('xlide_write_module');
        expect(named).toContain('xlide_run_tests');
        expect(named.filter((name) => /[A-Z]/.test(name))).toEqual([]);
    });

    it('says a module has no path on disk before it offers either way in', () => {
        // Measured in a real VS Code: a module document's fsPath answers
        // ENOENT to stat, read and write, and vscode.Uri.file of it matches no
        // open document, which is how the Claude Code extension's IDE calls
        // address a file. An agent that is not told this tries the path, or
        // edits an exported copy and reports the module changed.
        const without = AGENT_INSTRUCTIONS.slice(AGENT_INSTRUCTIONS.indexOf('## Without XLIDE\'s tools'));
        const steps = AGENT_INSTRUCTIONS_STEPS.join(' ');

        expect(without).toContain('no path on disk');
        expect(without.indexOf('does not exist')).toBeLessThan(without.indexOf('Export All Modules to Folder'));
        expect(without).toContain('Never report a module as changed because you wrote an exported');
        expect(steps).toMatch(/such as Claude and ChatGPT, cannot reach a module/);
    });

    it('is plain ASCII with no trailing spaces or tabs', () => {
        for (const text of [AGENT_INSTRUCTIONS, ...AGENT_INSTRUCTIONS_STEPS]) {
            expect([...text].filter((c) => c.charCodeAt(0) > 126 || (c.charCodeAt(0) < 32 && c !== '\n'))).toEqual([]);
            expect(text).not.toMatch(/ \n| $/);
        }
    });

    it('tells the person where each kind of agent reads its instructions', () => {
        const steps = AGENT_INSTRUCTIONS_STEPS.join('\n');

        for (const file of ['`.github/copilot-instructions.md`', '`CLAUDE.md`', '`AGENTS.md`']) {
            expect(steps).toContain(file);
        }
    });
});

describe('the Agent Instructions dialog', () => {
    it('opens from the Agentic AI section, between Welcome and Project Actions', () => {
        const html = renderXlideSidebarHtml(buildXlideSidebarModel({}));
        const at = (label: string) => html.indexOf(`<section class="section" aria-label="${label}">`);

        expect(at('Welcome')).toBeGreaterThanOrEqual(0);
        expect(at('Welcome')).toBeLessThan(at('Agentic AI'));
        expect(at('Agentic AI')).toBeLessThan(html.indexOf('id="agent-instructions-dialog"'));
        expect(html.indexOf('id="agent-instructions-dialog"')).toBeLessThan(at('Project Actions'));
        expect(html).toMatch(/<button class="actionCard secondary" type="button" data-dialog-open="agent-instructions-dialog" aria-haspopup="dialog"[^>]*>\s*<div class="label">Agent Instructions<\/div>/);
    });

    it('shows the steps, then the text read-only and exactly, with Copy', () => {
        const html = renderXlideSidebarHtml(buildXlideSidebarModel({}));
        const dialog = /<div class="dialogBackdrop" id="agent-instructions-dialog" data-dialog hidden>[\s\S]*?<\/div>\s*<\/div>\s*<\/div>/.exec(html)?.[0] ?? '';
        const text = /<textarea class="agentText" id="agent-instructions-text" readonly spellcheck="false">([\s\S]*?)<\/textarea>/.exec(dialog)?.[1];

        expect(dialog).toContain('role="dialog" aria-modal="true"');
        expect([...dialog.matchAll(/<li>/g)]).toHaveLength(AGENT_INSTRUCTIONS_STEPS.length);
        expect(dialog).toContain('<code>CLAUDE.md</code>');
        expect(text).toBeDefined();
        expect(decodeHtml(text!)).toBe(AGENT_INSTRUCTIONS);
        expect(dialog).toMatch(/<button type="button" data-agent-copy data-dialog-focus>Copy<\/button>/);
        expect(dialog).toContain('role="status"');
    });

    function openSidebar() {
        let receive: ((message: unknown) => void) | undefined;
        const posted: unknown[] = [];
        const view = {
            webview: {
                options: {},
                html: '',
                onDidReceiveMessage: (handler: (message: unknown) => void) => {
                    receive = handler;
                    return { dispose() { /* nothing held */ } };
                },
                postMessage: vi.fn(async (message: unknown) => {
                    posted.push(message);
                    return true;
                }),
            },
        };
        new XlideSidebarProvider({}).resolveWebviewView(view as never);
        return { send: (message: unknown) => receive!(message), posted };
    }

    it('copies the host\'s own text, whatever the webview sends, and says it did', async () => {
        const writeText = vi.mocked(vscode.env.clipboard.writeText);
        writeText.mockClear();
        const { send, posted } = openSidebar();

        send({ type: 'copyAgentInstructions', text: 'something else' });

        await vi.waitFor(() => expect(posted).toContainEqual({ type: 'agentInstructionsCopied' }));
        expect(writeText).toHaveBeenCalledTimes(1);
        expect(writeText).toHaveBeenCalledWith(AGENT_INSTRUCTIONS);
    });

    it('says so when the clipboard refuses, so the person can copy by hand', async () => {
        const writeText = vi.mocked(vscode.env.clipboard.writeText);
        writeText.mockClear();
        writeText.mockRejectedValueOnce(new Error('clipboard unavailable'));
        const { send, posted } = openSidebar();

        send({ type: 'copyAgentInstructions' });

        await vi.waitFor(() => expect(posted).toContainEqual({ type: 'agentInstructionsCopyFailed' }));
    });
});
