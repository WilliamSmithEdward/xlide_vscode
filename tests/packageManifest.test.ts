import { readFileSync, readdirSync } from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';

// Cross-consistency between package.json contributions and the source that
// must back them. Each failure mode ships silently without this: a declared
// command with no registration is a dead menu item, a menu entry naming an
// undeclared command never renders at all, and a language-model tool whose
// declaration and registration drift breaks the agent surface with no error.

interface Manifest {
    contributes: {
        commands: Array<{ command: string; title: string }>;
        menus: Record<string, Array<{ command?: string }>>;
        languageModelTools: Array<{ name: string }>;
        viewsWelcome?: Array<{ contents: string }>;
    };
}

const manifest = JSON.parse(readFileSync('package.json', 'utf8')) as Manifest;

function allSource(): string {
    const parts: string[] = [];
    const walk = (dir: string): void => {
        for (const entry of readdirSync(dir, { withFileTypes: true })) {
            const p = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                walk(p);
            } else if (entry.name.endsWith('.ts')) {
                parts.push(readFileSync(p, 'utf8'));
            }
        }
    };
    walk('src');
    return parts.join('\n');
}

const source = allSource();

function matches(pattern: RegExp): Set<string> {
    const found = new Set<string>();
    for (const m of source.matchAll(pattern)) {
        found.add(m[1]);
    }
    return found;
}

const registeredCommands = matches(/register(?:Xlide)?Command\(\s*['"]([^'"]+)['"]/g);
const registeredTools = matches(/registerTool[^(]*\(\s*['"]([^'"]+)['"]/g);
const declaredCommands = new Set(manifest.contributes.commands.map((c) => c.command));
const declaredTools = new Set(manifest.contributes.languageModelTools.map((t) => t.name));

/** Registered on purpose without a declaration: keybinding handlers and
 * commands attached directly to tree items, none of which render in menus
 * or the palette. Adding a command here is a deliberate decision. */
const INTERNAL_COMMANDS = new Set([
    'xlide.retryExplorerLoad',
    'xlide.vba.smartBackspace',
    'xlide.vba.smartTab',
    'xlide.vba.leaveSnippetAndCursorMove',
]);

describe('package manifest consistency', () => {
    it('every menu entry names a declared command', () => {
        const missing: string[] = [];
        for (const [menu, entries] of Object.entries(manifest.contributes.menus)) {
            for (const entry of entries) {
                if (entry.command && !declaredCommands.has(entry.command)) {
                    missing.push(`${menu}: ${entry.command}`);
                }
            }
        }
        expect(missing).toEqual([]);
    });

    it('every declared command is registered in source', () => {
        const dead = [...declaredCommands].filter((c) => !registeredCommands.has(c));
        expect(dead).toEqual([]);
    });

    it('every registered xlide command is declared or deliberately internal', () => {
        const undeclared = [...registeredCommands]
            .filter((c) => c.startsWith('xlide.'))
            .filter((c) => !declaredCommands.has(c) && !INTERNAL_COMMANDS.has(c));
        expect(undeclared).toEqual([]);
    });

    it('welcome-view command links name declared commands', () => {
        const missing: string[] = [];
        for (const view of manifest.contributes.viewsWelcome ?? []) {
            for (const m of view.contents.matchAll(/command:([A-Za-z0-9_.]+)/g)) {
                if (!declaredCommands.has(m[1])) {
                    missing.push(m[1]);
                }
            }
        }
        expect(missing).toEqual([]);
    });

    it('declared language-model tools and registrations match exactly', () => {
        expect([...declaredTools].filter((t) => !registeredTools.has(t))).toEqual([]);
        expect([...registeredTools].filter((t) => !declaredTools.has(t))).toEqual([]);
    });
});
