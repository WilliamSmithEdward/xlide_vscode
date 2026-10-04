import * as assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { BACKSPACE_NEEDS_EXTENSION_CONTEXT } from '../vbaEditorCommands';
import { until, workspaceRoot } from './support';

// Opt-in renderer probe. A separate process sends the physical key through
// Chromium's input API while this test deliberately occupies the extension
// host. Provider commands issued by the host cannot establish this property.
const rendererProbe = String.raw`
const fs = require('node:fs');
const [port, readyFile, busyFile, mode] = process.argv.slice(1);
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
(async () => {
    const targets = await (await fetch('http://127.0.0.1:' + port + '/json/list')).json();
    const target = targets.find(item => item.type === 'page' && item.url.includes('workbench'));
    if (!target) throw new Error('Owned integration workbench debugger target missing');
    const socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {
        socket.addEventListener('open', resolve, { once: true });
        socket.addEventListener('error', reject, { once: true });
    });
    let sequence = 0;
    const pending = new Map();
    socket.addEventListener('message', event => {
        const message = JSON.parse(event.data);
        const entry = pending.get(message.id);
        if (!entry) return;
        pending.delete(message.id);
        clearTimeout(entry.timer);
        if (message.error) entry.reject(new Error(message.error.message));
        else entry.resolve(message.result);
    });
    const call = (method, params = {}) => new Promise((resolve, reject) => {
        const id = ++sequence;
        const timer = setTimeout(() => { pending.delete(id); reject(new Error('Debugger request timed out: ' + method)); }, 5000);
        pending.set(id, { resolve, reject, timer });
        socket.send(JSON.stringify({ id, method, params }));
    });
    const evaluate = async expression => {
        const response = await call('Runtime.evaluate', { expression, returnByValue: true });
        if (response.exceptionDetails) throw new Error('Renderer expression failed');
        return response.result.value;
    };
    const readLine = "Array.from(document.querySelectorAll('.monaco-editor .view-line')).map(line => line.textContent).find(text => text.includes('ThisWorkbook.Sheets(1).ce'))";
    if (!(await evaluate(readLine))?.endsWith('.cez')) throw new Error('Synthetic test line is not visible');
    await evaluate("(() => { const input = document.querySelector('.monaco-editor.focused .inputarea') || document.querySelector('.monaco-editor .inputarea'); if (input) input.focus(); return document.activeElement?.className; })()");
    fs.writeFileSync(readyFile, 'ready');
    const deadline = Date.now() + 15000;
    while (!fs.existsSync(busyFile)) {
        if (Date.now() > deadline) throw new Error('Busy-host signal did not arrive');
        await delay(10);
    }
    const { busyUntil } = JSON.parse(fs.readFileSync(busyFile, 'utf8'));
    const deleteKey = async () => {
        await call('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Backspace', code: 'Backspace', windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 8 });
        await call('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Backspace', code: 'Backspace', windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 8 });
    };
    if (mode === 'stress') {
        const samples = [];
        const until = async (check, phase) => {
            const deadline = Date.now() + 4000;
            while (!(await check())) {
                if (Date.now() > deadline) throw new Error('Expected renderer update did not paint: ' + phase);
                await delay(5);
            }
        };
        for (let i = 0; i < 24; i++) {
            const idleMs = [0, 30, 250, 350][i % 4];
            await delay(idleMs);
            const before = Date.now();
            await deleteKey();
            await until(async () => (await evaluate(readLine))?.endsWith('.ce'), 'Backspace cycle ' + i);
            const deleted = Date.now();
            await until(async () => await evaluate("Array.from(document.querySelectorAll('.suggest-widget.visible .monaco-list-row')).some(row => row.checkVisibility({ visibilityProperty: true, opacityProperty: true }) && row.textContent.includes('Cells'))"), 'menu recovery cycle ' + i);
            const menuPaintMs = Date.now() - deleted;
            const typed = Date.now();
            await call('Input.insertText', { text: 'z' });
            await until(async () => (await evaluate(readLine))?.endsWith('.cez'), 'typing cycle ' + i);
            const typingPaintMs = Date.now() - typed;
            const missed = Date.now();
            // Do not count a stale Cells row from the preceding .ce cycle as
            // successful recovery. The .cez miss must clear it first.
            await until(async () => !(await evaluate("Array.from(document.querySelectorAll('.suggest-widget.visible .monaco-list-row')).some(row => row.checkVisibility({ visibilityProperty: true, opacityProperty: true }) && row.textContent.includes('Cells'))")), 'miss menu invalidation cycle ' + i);
            samples.push({ idleMs, backspacePaintMs: deleted - before, menuPaintMs, typingPaintMs, missClearMs: Date.now() - missed });
        }
        console.log(JSON.stringify({ samples }));
        socket.close();
        return;
    }
    const started = Date.now();
    await deleteKey();
    let deletedAt;
    const observationEnd = Math.min(started + 600, busyUntil - 100);
    while (Date.now() < observationEnd) {
        if ((await evaluate(readLine))?.endsWith('.ce')) { deletedAt = Date.now(); break; }
        await delay(10);
    }
    console.log(JSON.stringify({ deletedWhileBusy: deletedAt !== undefined && deletedAt < busyUntil, elapsedMs: (deletedAt || Date.now()) - started }));
    socket.close();
})().catch(error => { console.error(error.message); process.exit(1); });
`;

export interface RendererBackspaceResult {
    deletedWhileBusy?: boolean;
    elapsedMs?: number;
    samples?: { idleMs: number; backspacePaintMs: number; menuPaintMs: number; typingPaintMs: number; missClearMs: number }[];
}

export async function runRendererBackspaceProbe(mode: 'busy' | 'stress', routedThroughHost = false): Promise<RendererBackspaceResult> {
    const port = Number(process.env.XLIDE_UI_DEBUG_PORT);
    assert.ok(Number.isInteger(port) && port >= 1024 && port <= 65535, 'an owned integration renderer debugger port is required');
    const root = workspaceRoot();
    const readyFile = path.join(root, 'backspace-renderer.ready');
    const busyFile = path.join(root, 'backspace-renderer.busy');
    for (const marker of [readyFile, busyFile]) { fs.rmSync(marker, { force: true }); }
    const child = spawn(process.execPath, ['-e', rendererProbe, String(port), readyFile, busyFile, mode], {
        env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '', stderr = '';
    child.stdout.on('data', data => { stdout += String(data); });
    child.stderr.on('data', data => { stderr += String(data); });
    const finished = new Promise<number | null>((resolve, reject) => {
        child.once('error', reject); child.once('close', resolve);
    });
    void finished.catch(() => {});
    try {
        await until(() => {
            assert.equal(stderr, '', stderr);
            return fs.existsSync(readyFile) || undefined;
        }, 'renderer input probe should become ready', 12000);
        if (routedThroughHost) {
            // Positive control models the former always-bound keybinding.
            await vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, true);
        }
        const busyUntil = Date.now() + (mode === 'busy' ? 1200 : 0);
        fs.writeFileSync(busyFile, JSON.stringify({ busyUntil }));
        if (mode === 'busy') {
            // Deliberate test-only stall. The separate renderer must continue
            // deleting ordinary code throughout this occupied-host interval.
            while (Date.now() < busyUntil) { /* occupy the extension-host event loop */ }
        }
        assert.equal(await finished, 0, stderr);
        return JSON.parse(stdout.trim()) as RendererBackspaceResult;
    } finally {
        child.kill();
        for (const marker of [readyFile, busyFile]) { fs.rmSync(marker, { force: true }); }
    }
}
