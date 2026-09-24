import { readFileSync } from 'fs';
import { afterEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    window: { activeTextEditor: undefined, visibleTextEditors: [] },
    workspace: { workspaceFolders: [] },
}));

import * as vscode from 'vscode';
import {
    AGENT_INSTRUCTIONS,
    AGENT_INSTRUCTIONS_STEPS,
    MCP_EDIT_MIRROR_HINT,
    MCP_EDIT_MIRROR_LABEL,
} from '../src/agentInstructions';
import { fakeConfig } from './helpers/fakeConfig';
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

    it('tells an agent on the MCP server what mirroring shows, so it can tell the user to turn it on', () => {
        // The server's tool results say whether an XLIDE window took each
        // edit. Without this an agent whose edits reach no tree has nothing
        // to tell the user, who then sees no diff and no Revert.
        const server = AGENT_INSTRUCTIONS.slice(AGENT_INSTRUCTIONS.indexOf('## The XLIDE MCP server'));

        expect(server).toContain(`"${MCP_EDIT_MIRROR_LABEL}"`);
        expect(server).toContain('Agent Instructions dialog');
        expect(server).toContain('xlide_vscode');
        expect(server).toMatch(/notified is false[^.]*no XLIDE window took the report/);
        expect(server).toContain('turn that toggle on');
        expect(server).toContain('review is "pending"');
    });

    it('is plain ASCII with no trailing spaces or tabs', () => {
        for (const text of [AGENT_INSTRUCTIONS, ...AGENT_INSTRUCTIONS_STEPS, MCP_EDIT_MIRROR_LABEL, MCP_EDIT_MIRROR_HINT]) {
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
    afterEach(() => {
        // Back to the mock's own settings, which read every value as its default.
        vi.mocked(vscode.workspace.getConfiguration).mockReset();
    });

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

    it('offers the mirroring toggle under the steps, showing the setting as it stands', () => {
        const model = buildXlideSidebarModel({});
        const on = renderXlideSidebarHtml(model, { mcpEditMirror: true });
        const off = renderXlideSidebarHtml(model, { mcpEditMirror: false });
        const toggle = /<input type="checkbox" data-mcp-mirror aria-describedby="agent-mcp-mirror-hint"( checked)?>/;

        expect(toggle.exec(on)?.[1]).toBe(' checked');
        expect(toggle.exec(off)?.[1]).toBeUndefined();
        expect(on).toContain(`<span>${MCP_EDIT_MIRROR_LABEL}</span>`);
        expect(on).toContain(`id="agent-mcp-mirror-hint">${MCP_EDIT_MIRROR_HINT.replace(/'/g, '&#39;')}</p>`);
        expect(on.indexOf('id="agent-instructions-steps"')).toBeLessThan(on.indexOf('data-mcp-mirror'));
        expect(on.indexOf('data-mcp-mirror')).toBeLessThan(on.indexOf('for="agent-instructions-text"'));
    });

    it('leaves the toggle out where the build does not mirror', () => {
        // The browser build: no port to listen on, so nothing to turn on.
        const html = renderXlideSidebarHtml(buildXlideSidebarModel({}));

        expect(html).not.toContain('<input type="checkbox" data-mcp-mirror');
        expect(html).not.toContain('id="agent-mcp-mirror-hint"');
    });

    function openSidebar(options: ConstructorParameters<typeof XlideSidebarProvider>[0] = {}) {
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
        const provider = new XlideSidebarProvider(options);
        provider.resolveWebviewView(view as never);
        return { send: (message: unknown) => receive!(message), posted, provider };
    }

    it('writes the mirroring setting when the toggle changes', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        vi.mocked(vscode.workspace.getConfiguration).mockReturnValue(fakeConfig({}, new Set(), updates));
        const { send } = openSidebar({ offersMcpEditMirror: true });

        send({ type: 'setMcpEditMirror', enabled: false });

        await vi.waitFor(() => expect(updates).toEqual([{ key: 'agent.mirrorMcpEdits', value: false, target: true }]));
    });

    it('writes nothing where the build does not mirror, or for a message that is not a yes or no', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        vi.mocked(vscode.workspace.getConfiguration).mockReturnValue(fakeConfig({}, new Set(), updates));

        openSidebar().send({ type: 'setMcpEditMirror', enabled: false });
        openSidebar({ offersMcpEditMirror: true }).send({ type: 'setMcpEditMirror', enabled: 'off' });
        await new Promise((resolve) => setTimeout(resolve, 20));

        expect(updates).toEqual([]);
    });

    it('shows the setting as it stands when writing it failed, and says so', async () => {
        const config = fakeConfig({ 'agent.mirrorMcpEdits': true }, new Set(['agent.mirrorMcpEdits']));
        config.update = () => Promise.reject(new Error('settings file is read-only'));
        vi.mocked(vscode.workspace.getConfiguration).mockReturnValue(config);
        const { send, posted } = openSidebar({ offersMcpEditMirror: true });

        send({ type: 'setMcpEditMirror', enabled: false });

        await vi.waitFor(() => expect(posted).toContainEqual({ type: 'mcpEditMirror', enabled: true, failed: true }));
    });

    it('tells the page when the setting changed anywhere else', async () => {
        vi.mocked(vscode.workspace.getConfiguration).mockReturnValue(fakeConfig({ 'agent.mirrorMcpEdits': false }));
        const { provider, posted } = openSidebar({ offersMcpEditMirror: true });

        await provider.postMcpEditMirror();

        expect(posted).toContainEqual({ type: 'mcpEditMirror', enabled: false, failed: false });
    });

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
