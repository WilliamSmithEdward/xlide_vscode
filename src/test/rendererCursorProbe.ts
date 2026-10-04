import * as assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { until, workspaceRoot } from './support';

export type CursorDirection = 'up' | 'down' | 'left' | 'right';
interface CursorState {
    placeholders: number;
    widgets: number;
    beforeTyped: boolean;
    middleTyped: boolean;
}
export interface RendererCursorResult {
    movedWhileBusy?: boolean;
    typedWhileBusy?: boolean;
    elapsedMs: number;
    before: CursorState;
    after: CursorState;
}

// An independent process observes the renderer while the extension host is
// deliberately occupied. Provider command timings cannot prove native routing.
const renderer = String.raw`
const fs = require('node:fs');
const [port, ready, busy, direction, typeAfter] = process.argv.slice(1);
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
(async () => {
 const targets = await (await fetch('http://127.0.0.1:' + port + '/json/list')).json();
 const target = targets.find(t => t.type === 'page' && t.url.includes('workbench'));
 if (!target) throw new Error('Owned snippet workbench missing');
 const socket = new WebSocket(target.webSocketDebuggerUrl);
 await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, {once: true}); socket.addEventListener('error', reject, {once: true}); });
 let sequence = 0; const pending = new Map();
 socket.addEventListener('message', event => { const msg = JSON.parse(event.data); const item = pending.get(msg.id); if (!item) return; pending.delete(msg.id); clearTimeout(item.timer); if (msg.error) item.reject(new Error(msg.error.message)); else item.resolve(msg.result); });
 const call = (method, params = {}) => new Promise((resolve, reject) => { const id = ++sequence; const timer = setTimeout(() => { pending.delete(id); reject(new Error('Debugger request timed out: ' + method)); }, 5000); pending.set(id, {resolve, reject, timer}); socket.send(JSON.stringify({id, method, params})); });
 const stateExpression = "(() => { const editor = document.querySelector('.monaco-editor.focused'); const cursor = Array.from(editor?.querySelectorAll('.cursors-layer .cursor') ?? []).find(n => n.getBoundingClientRect().height > 0); const caret = cursor?.getBoundingClientRect(); const rows = Array.from(editor?.querySelectorAll('.view-line') ?? []); const row = rows.find(n => { const b = n.getBoundingClientRect(); return caret && caret.top >= b.top && caret.top < b.bottom; }); return { top: caret?.top, left: caret?.left, middle: row?.textContent.includes('MiddleMarker'), placeholders: editor?.querySelectorAll('.snippet-placeholder').length, widgets: Array.from(document.querySelectorAll('.suggest-widget.visible')).filter(n => n.checkVisibility({visibilityProperty:true,opacityProperty:true})).length, beforeTyped: rows.some(node => node.textContent.includes('BeforezLine')), middleTyped: rows.some(node => node.textContent.includes('MiddlezMarker')), focused: document.hasFocus() }; })()";
 const read = async () => { const response = await call('Runtime.evaluate', {expression: stateExpression, returnByValue: true}); if (response.exceptionDetails) throw new Error('Caret expression failed'); return response.result.value; };
 let before; const prepareEnd = Date.now() + 10000;
 do { before = await read(); if (Date.now() > prepareEnd) throw new Error('Snippet fixture caret not ready: ' + JSON.stringify(before)); if (!before.middle || before.widgets) await sleep(10); } while (!before.middle || before.widgets);
 fs.writeFileSync(ready, JSON.stringify(before));
 const signalEnd = Date.now() + 10000;
 while (!fs.existsSync(busy)) { if (Date.now() > signalEnd) throw new Error('Busy snippet signal missing'); await sleep(5); }
 const busyUntil = JSON.parse(fs.readFileSync(busy, 'utf8')).busyUntil;
 const names = {up:['ArrowUp',38],down:['ArrowDown',40],left:['ArrowLeft',37],right:['ArrowRight',39]};
 const [key, virtualKey] = names[direction]; const started = Date.now();
 for (const type of ['keyDown','keyUp']) await call('Input.dispatchKeyEvent', {type, key, code:key, windowsVirtualKeyCode:virtualKey, nativeVirtualKeyCode:virtualKey});
 if (typeAfter === '1') await call('Input.insertText', {text:'z'});
 let after, movedAt; const end = Math.min(started + 600, busyUntil - 100);
 while (Date.now() < end) { after = await read(); const vertical = direction === 'up' || direction === 'down'; if (typeAfter === '1' ? after.beforeTyped : vertical ? after.top !== before.top : after.left !== before.left) { movedAt = Date.now(); break; } await sleep(5); }
 console.log(JSON.stringify({direction,...(typeAfter === '1' ? {typedWhileBusy:movedAt !== undefined && movedAt < busyUntil} : {movedWhileBusy:movedAt !== undefined && movedAt < busyUntil}),elapsedMs:(movedAt || Date.now())-started,before,after}));
 socket.close();
})().catch(error => { console.error(error.message); process.exit(1); });
`;

export async function runRendererCursorProbe(direction: CursorDirection, snippet: boolean, typeAfter = false): Promise<RendererCursorResult> {
    const port = Number(process.env.XLIDE_UI_DEBUG_PORT);
    assert.ok(Number.isInteger(port) && port >= 1024 && port <= 65535, 'owned renderer debugger port required');
    const root = workspaceRoot();
    const ready = path.join(root, 'snippet-renderer.ready');
    const busy = path.join(root, 'snippet-renderer.busy');
    for (const marker of [ready, busy]) { fs.rmSync(marker, { force: true }); }
    const child = spawn(process.execPath, ['-e', renderer, String(port), ready, busy, direction, typeAfter ? '1' : '0'], {
        env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
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
            return fs.existsSync(ready) || undefined;
        }, 'renderer cursor probe should become ready', 12000);
        const before = JSON.parse(fs.readFileSync(ready, 'utf8')) as CursorState;
        assert.equal(before.widgets, 0);
        assert.equal(before.placeholders > 0, snippet, 'probe must exercise the real snippet state');
        const busyUntil = Date.now() + 1200;
        fs.writeFileSync(busy, JSON.stringify({ busyUntil }));
        while (Date.now() < busyUntil) { /* test-only occupied host */ }
        assert.equal(await finished, 0, stderr);
        return JSON.parse(stdout.trim()) as RendererCursorResult;
    } finally {
        child.kill();
        for (const marker of [ready, busy]) { fs.rmSync(marker, { force: true }); }
    }
}
