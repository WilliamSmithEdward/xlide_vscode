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
    if (mode !== 'transition' && mode !== 'cleanup' && !(await evaluate(readLine))?.endsWith('.cez')) throw new Error('Synthetic test line is not visible');
    await evaluate("(() => { const input = document.querySelector('.monaco-editor.focused .inputarea') || document.querySelector('.monaco-editor .inputarea'); if (input) input.focus(); return document.activeElement?.className; })()");
    fs.writeFileSync(readyFile, 'ready');
    const deadline = Date.now() + 15000;
    while (!fs.existsSync(busyFile)) {
        if (Date.now() > deadline) throw new Error('Busy-host signal did not arrive');
        await delay(10);
    }
    const { busyUntil, cleanup, stress } = JSON.parse(fs.readFileSync(busyFile, 'utf8'));
    const pressKey = async (key, code, virtualKey, modifiers = 0) => {
        await call('Input.dispatchKeyEvent', { type: 'keyDown', key, code, modifiers, windowsVirtualKeyCode: virtualKey, nativeVirtualKeyCode: virtualKey });
        await call('Input.dispatchKeyEvent', { type: 'keyUp', key, code, modifiers, windowsVirtualKeyCode: virtualKey, nativeVirtualKeyCode: virtualKey });
    };
    const deleteKey = () => pressKey('Backspace', 'Backspace', 8);
    if (mode === 'cleanup') {
        const readRow = "Array.from(document.querySelectorAll('.monaco-editor.focused .view-line')).map(row => row.textContent.replace(/\\u00a0/g, ' ').replace(/\\u200b/g, ''))[" + cleanup.line + "]";
        if ((await evaluate(readRow)) !== cleanup.before) throw new Error('Cleanup fixture is not visible');
        const samples = [];
        for (const expected of cleanup.after) {
            const started = Date.now();
            await deleteKey();
            const deadline = Date.now() + 4000;
            for (;;) {
                const visible = await evaluate(readRow);
                if (visible === expected || (expected === '' && visible === ' ')) break;
                if (Date.now() > deadline) throw new Error('Smart cleanup did not restore its expected visible line');
                await delay(5);
            }
            samples.push(Date.now() - started);
        }
        console.log(JSON.stringify({ cleanupPaintMs: samples }));
        socket.close(); return;
    }
    if (mode === 'stress') {
        const samples = [];
        const profileEnabled = process.env.XLIDE_PERF_CPU_PROFILE === '1';
        const profileStartRequestedAt = Date.now();
        if (profileEnabled) {
            await call('Profiler.enable');
            await call('Profiler.setSamplingInterval', { interval: 1000 });
            await call('Profiler.start');
        }
        const profileStartedAt = Date.now();
        let expectedNonce;
        const navigationSamples = [];
        const captureNavigation = async (cycle, phase) => {
            if (process.env.XLIDE_PERF_NAV_DIAGNOSTICS !== '1') return;
            const state = await evaluate("(() => { const editor = document.querySelector('.monaco-editor.focused'); const rows = Array.from(editor?.querySelectorAll('.view-line') ?? []); const cursor = Array.from(editor?.querySelectorAll('.cursors-layer .cursor') ?? []).find(node => node.getBoundingClientRect().height > 0); const caret = cursor?.getBoundingClientRect(); const row = rows.find(node => { const box = node.getBoundingClientRect(); return caret && caret.top >= box.top && caret.top < box.bottom; }); return { documentFocused: document.hasFocus(), focusedEditors: document.querySelectorAll('.monaco-editor.focused').length, caretOnMember: row?.textContent.includes('ThisWorkbook.Sheets(1).'), caretOnNonce: row?.textContent.includes('LatencyValue'), caretRowLength: row?.textContent.length, visibleWidgets: Array.from(document.querySelectorAll('.suggest-widget.visible')).filter(node => node.checkVisibility({ visibilityProperty: true, opacityProperty: true })).length, memberRows: rows.filter(node => node.textContent.includes('ThisWorkbook.Sheets(1).')).length }; })()");
            navigationSamples.push({ cycle, phase, ...state });
            if (navigationSamples.length > 24) navigationSamples.shift();
        };
        const until = async (check, phase) => {
            const deadline = Date.now() + 4000;
            while (!(await check())) {
                if (Date.now() > deadline) {
                    const widgets = await evaluate("Array.from(document.querySelectorAll('.suggest-widget.visible')).map(node => ({ shown: node.checkVisibility({ visibilityProperty: true, opacityProperty: true }), classes: node.className, message: node.querySelector('.message')?.textContent, rows: node.querySelectorAll('.monaco-list-row').length, kinds: Array.from(node.querySelectorAll('.monaco-list-row')).map(row => row.querySelector('.suggest-icon')?.className) }))");
                    const state = await evaluate("((expectedNonce) => { const input = document.activeElement; const rows = Array.from(document.querySelectorAll('.monaco-editor .view-line')); const row = rows.find(row => row.textContent.includes('ThisWorkbook.Sheets(1).')); return { focused: input?.className, documentFocused: document.hasFocus(), visibility: document.visibilityState, focusedEditors: document.querySelectorAll('.monaco-editor.focused').length, endsCe: row?.textContent.endsWith('.ce'), endsCez: row?.textContent.endsWith('.cez'), syntheticLength: row?.textContent.length, nonceRows: rows.filter(row => row.textContent.includes('LatencyValue')).map(row => ({ length: row.textContent.length, endsExpected: expectedNonce ? row.textContent.endsWith(expectedNonce) : undefined, normalizedEndsExpected: expectedNonce ? row.textContent.replace(/\\u00a0/g, ' ').replace(/\\u200b/g, '').endsWith(expectedNonce) : undefined, ghostNodes: row.querySelectorAll('.ghost-text').length })) }; })(" + JSON.stringify(expectedNonce ?? null) + ")");
                    throw new Error('Expected renderer update did not paint: ' + phase + '; widgets=' + JSON.stringify(widgets) + '; state=' + JSON.stringify(state) + '; navigation=' + JSON.stringify(navigationSamples));
                }
                await delay(5);
            }
        };
        const visibleHover = "Array.from(document.querySelectorAll('.monaco-hover')).filter(node => node.checkVisibility({ visibilityProperty: true, opacityProperty: true })).map(node => node.textContent).join(' ')";
        const hoverSamples = [];
        const readHoverPoint = () => evaluate("(() => { for (const row of document.querySelectorAll('.monaco-editor.focused .view-line')) { const walker = document.createTreeWalker(row, NodeFilter.SHOW_TEXT); let text; while ((text = walker.nextNode())) { const start = text.textContent.indexOf('LatencyValue'); if (start < 0) continue; const range = document.createRange(); range.setStart(text, start + 2); range.setEnd(text, start + 3); const box = range.getBoundingClientRect(); return { x: box.x + box.width / 2, y: box.y + box.height / 2 }; } } })()");
        const showHover = async () => {
            // Navigation can leave the viewport scrolled. Clear the list and
            // reveal the short member line before testing a mouse hit above it.
            await pressKey('Escape', 'Escape', 27);
            await pressKey('Home', 'Home', 36);
            await pressKey('End', 'End', 35);
            await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 1, y: 1 });
            await until(async () => !(await evaluate(visibleHover)).includes('LatencyValue'), 'old hover dismissal');
            const viewport = await evaluate("(() => { const rect = document.querySelector('.monaco-editor.focused .editor-scrollable').getBoundingClientRect(); return { left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom }; })()");
            let point;
            await until(async () => {
                point = await readHoverPoint();
                return point && point.x >= viewport.left && point.x <= viewport.right &&
                    point.y >= viewport.top && point.y <= viewport.bottom;
            }, 'hover target viewport visibility');
            const started = Date.now();
            await call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point });
            await until(async () => (await evaluate(visibleHover)).includes('LatencyValue As Long'), 'resolved mouse hover');
            hoverSamples.push(Date.now() - started);
            await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 1, y: 1 });
            await until(async () => !(await evaluate(visibleHover)).includes('LatencyValue'), 'resolved hover dismissal');
        };
        try {
            for (let i = 0; i < stress.cycles; i++) {
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
                // Other providers can have matching rows. They must not leave a
                // loading/empty message after the keyboard-driven miss settles.
                await until(async () => !(await evaluate("Array.from(document.querySelectorAll('.suggest-widget.visible.message')).some(node => node.checkVisibility({ visibilityProperty: true, opacityProperty: true }))")), 'miss status-message dismissal cycle ' + i);
                if (stress.assertMissHidden) {
                    await until(async () => !(await evaluate("Array.from(document.querySelectorAll('.suggest-widget.visible')).some(node => node.checkVisibility({ visibilityProperty: true, opacityProperty: true }))")), 'miss widget dismissal cycle ' + i);
                }
                samples.push({ ...(profileEnabled ? { startedAt: before } : {}), idleMs, backspacePaintMs: deleted - before, menuPaintMs, typingPaintMs, missClearMs: Date.now() - missed });
                if (stress.hover && (i + 1) % 16 === 0) { await showHover(); }
                if ((i + 1) % 100 === 0) fs.writeFileSync(readyFile, JSON.stringify({ completed: i + 1, cycles: stress.cycles }));
                if (stress.freshSources) {
                    // Change only the synthetic preceding statement, keeping the
                    // member expression identical while defeating source-text reuse.
                    await captureNavigation(i, 'before fresh edit');
                    await pressKey('Escape', 'Escape', 27);
                    await captureNavigation(i, 'after Escape');
                    await pressKey('ArrowUp', 'ArrowUp', 38);
                    await captureNavigation(i, 'after ArrowUp');
                    await pressKey('End', 'End', 35);
                    await captureNavigation(i, 'after End');
                    await pressKey('Home', 'Home', 36, 8); // select to first non-whitespace
                    await captureNavigation(i, 'after ShiftHome');
                    const nonce = " 'n" + i.toString(36).padStart(6, '0');
                    expectedNonce = nonce.trimStart();
                    await call('Input.insertText', { text: stress.nonceStatement + nonce });
                    await until(async () => await evaluate("Array.from(document.querySelectorAll('.monaco-editor.focused .view-line')).some(row => row.textContent.endsWith(" + JSON.stringify(nonce.trimStart()) + "))"), 'fresh synthetic source cycle ' + i);
                    await captureNavigation(i, 'after nonce');
                    await pressKey('ArrowDown', 'ArrowDown', 40);
                    await captureNavigation(i, 'after ArrowDown');
                    await pressKey('End', 'End', 35);
                    await captureNavigation(i, 'after final End');
                }
            }
        } finally {
            fs.writeFileSync(busyFile + '.renderer-observations.json', JSON.stringify({ samples, hoverSamples, navigationSamples }));
            if (profileEnabled) {
                const profileStopRequestedAt = Date.now();
                const stopped = await call('Profiler.stop');
                fs.writeFileSync(busyFile + '.renderer.cpuprofile', JSON.stringify(stopped.profile));
                fs.writeFileSync(busyFile + '.renderer-timings.json', JSON.stringify({ profileStartRequestedAt, profileStartedAt, profileStopRequestedAt, profileStoppedAt: Date.now(), samples, hoverSamples }));
            }
        }
        console.log(JSON.stringify({ samples, hoverSamples }));
        socket.close();
        return;
    }
    if (mode === 'transition') {
        await call('Input.insertText', { text: 'Debug.Print ThisWorkbook.Sheets(1).cez' });
        const typedDeadline = Date.now() + 500;
        while (!(await evaluate(readLine))?.endsWith('.cez')) {
            if (Date.now() > typedDeadline) throw new Error('Native typing did not reach the transition probe');
            await delay(5);
        }
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
    cleanupPaintMs?: number[];
    hoverSamples?: number[];
    samples?: { idleMs: number; backspacePaintMs: number; menuPaintMs: number; typingPaintMs: number; missClearMs: number }[];
}

export async function runRendererBackspaceProbe(mode: 'busy' | 'stress' | 'transition' | 'cleanup', staleCleanupContext = false, cleanup?: { line: number; before: string; after: string[] }, stress = { cycles: 24, hover: false, assertMissHidden: false, freshSources: false, nonceStatement: '' }): Promise<RendererBackspaceResult> {
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
        if (staleCleanupContext) {
            // A stale true flag must still delete in the renderer first.
            await vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, true);
        }
        const busyUntil = Date.now() + (mode === 'busy' || mode === 'transition' ? 1200 : 0);
        fs.writeFileSync(busyFile, JSON.stringify({ busyUntil, cleanup, stress }));
        if (mode === 'busy' || mode === 'transition') {
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
