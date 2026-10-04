import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { spawn } from 'node:child_process';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, workspaceRoot } from './support';

const probe = String.raw`
const fs = require('node:fs');
const [port, ready, go, busy, kind, move, typeAfter] = process.argv.slice(1);
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
(async () => {
 const targets = await (await fetch('http://127.0.0.1:' + port + '/json/list')).json();
 const target = targets.find(t => t.type === 'page' && t.url.includes('workbench'));
 if (!target) throw new Error('Owned Enter workbench missing');
 const socket = new WebSocket(target.webSocketDebuggerUrl);
 await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, {once:true}); socket.addEventListener('error', reject, {once:true}); });
 let sequence = 0; const pending = new Map();
 socket.addEventListener('message', event => { const msg = JSON.parse(event.data); const item = pending.get(msg.id); if (!item) return; pending.delete(msg.id); clearTimeout(item.timer); if (msg.error) item.reject(new Error(msg.error.message)); else item.resolve(msg.result); });
 const call = (method, params={}) => new Promise((resolve,reject) => { const id=++sequence; const timer=setTimeout(() => { pending.delete(id); reject(new Error('Debugger request timed out: '+method)); },5000); pending.set(id,{resolve,reject,timer}); socket.send(JSON.stringify({id,method,params})); });
 const expression = "(() => { const editor=document.querySelector('.monaco-editor.focused'); const cursor=Array.from(editor?.querySelectorAll('.cursors-layer .cursor') ?? []).find(n=>n.getBoundingClientRect().height>0); const rect=cursor?.getBoundingClientRect(); const rows=Array.from(editor?.querySelectorAll('.view-line') ?? []); const row=rows.find(n=>{const b=n.getBoundingClientRect();return rect && rect.top>=b.top && rect.top<b.bottom;}); return {top:rect?.top,left:rect?.left,text:row?.textContent.replaceAll('\\u00a0',' ').replaceAll('\\u200b',''),widgets:Array.from(document.querySelectorAll('.suggest-widget.visible')).filter(n=>n.checkVisibility({visibilityProperty:true,opacityProperty:true})).length,focused:document.hasFocus()}; })()";
 const read = async () => { const result=await call('Runtime.evaluate',{expression,returnByValue:true}); if(result.exceptionDetails) throw new Error('Caret expression failed'); return result.result.value; };
 const key = async (name,code) => { for(const type of ['keyDown','keyUp']) await call('Input.dispatchKeyEvent',{type,key:name,code:name==='Enter'?'Enter':name,windowsVirtualKeyCode:code,nativeVirtualKeyCode:code,...(name==='Enter' && type==='keyDown'?{text:'\r',unmodifiedText:'\r'}:{})}); };
 let initial;const prepareEnd=Date.now()+10000;
 do { initial=await read(); if(Date.now()>prepareEnd) throw new Error('Initial Enter caret missing: '+JSON.stringify(initial)); const expected=kind==='block'?/With ActiveSheet/i:kind==='comment'?/' note/:/\.Name =/i; if(!expected.test(initial.text ?? '') || initial.widgets) await sleep(10); } while(!(kind==='block'?/With ActiveSheet/i:kind==='comment'?/' note/:/\.Name =/i).test(initial.text ?? '') || initial.widgets);
 fs.writeFileSync(ready,JSON.stringify(initial));
 const goEnd=Date.now()+10000;while(!fs.existsSync(go)){if(Date.now()>goEnd)throw new Error('Enter signal missing');await sleep(5);}
 await key('Enter',13);
 const seedEnd=Date.now()+10000;let seeded;
 do { seeded=await read(); if(Date.now()>seedEnd)throw new Error('Seeded body/busy signal missing: '+JSON.stringify({initial,seeded,busy:fs.existsSync(busy)})); const text=seeded.text?.trim(); if(fs.existsSync(busy) && (kind==='comment'?text==="'":text==='.'))break; await sleep(5); } while(true);
 const busyUntil=JSON.parse(fs.readFileSync(busy,'utf8')).busyUntil;
 let navigated=seeded;let movedAt;
 if(move==='1'){
  const started=Date.now();await key('ArrowUp',38);
  const end=Math.min(started+600,busyUntil-50);
  do { navigated=await read(); if(navigated.top!==seeded.top){movedAt=Date.now();break;}await sleep(5); } while(Date.now()<end);
  if(movedAt===undefined)throw new Error('Native Up did not move during pending Enter acknowledgement: '+JSON.stringify({seeded,navigated}));
 }
 let typedWhileBusy;
 if(typeAfter==='1'){
  await call('Input.insertText',{text:'z'});
  const typingEnd=Math.min(Date.now()+600,busyUntil-50);
  do { navigated=await read();if(navigated.text?.includes('z')){typedWhileBusy=Date.now()<busyUntil;break;}await sleep(5); } while(Date.now()<typingEnd);
  if(!typedWhileBusy)throw new Error('Rapid typing did not reach the native-selected row while acknowledgement waited');
 }
 while(Date.now()<busyUntil+700)await sleep(10);
 const settled=await read();
 console.log(JSON.stringify({initial,seeded,navigated,settled,movedWhileBusy:movedAt!==undefined && movedAt<busyUntil,typedWhileBusy}));socket.close();
})().catch(error=>{console.error(error.message);process.exit(1);});
`;

suite('Native Enter caret ownership', () => {
    const port = Number(process.env.XLIDE_UI_DEBUG_PORT);
    setup(async function () {
        if (!Number.isInteger(port) || port < 1024 || port > 65535) { this.skip(); }
        await activate();
    });
    teardown(async () => {
        if (!port) { return; }
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
        await closeAllEditors();
    });
    for (const kind of ['block', 'comment', 'member'] as const) for (const action of ['stay', 'up', 'type'] as const) {
        const move = action !== 'stay';
        test(`${kind} ${action} retains caret ownership while acknowledgement waits`, async () => {
            const source = kind === 'block' ? 'Sub NativeEnterProbe()\n    AnchorMarker\n    With ActiveSheet\n    End With\nEnd Sub\n'
                : kind === 'comment' ? "Sub NativeEnterProbe()\n    AnchorMarker\n    ' note\nEnd Sub\n"
                : 'Sub NativeEnterProbe()\n    AnchorMarker\n    With ActiveSheet\n        .Name = "test"\n    End With\nEnd Sub\n';
            const file = path.join(workspaceRoot(), `NativeEnter-${kind}-${action}.bas`);
            fs.writeFileSync(file, source);
            await open(vscode.Uri.file(file));
            const editor = vscode.window.activeTextEditor!;
            const openerLine = kind === 'member' ? 3 : 2;
            const caret = new vscode.Position(openerLine, editor.document.lineAt(openerLine).text.length);
            editor.selection = new vscode.Selection(caret, caret);
            await vscode.commands.executeCommand('hideSuggestWidget');
            await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
            const ready = path.join(workspaceRoot(), 'enter-probe.ready');
            const go = path.join(workspaceRoot(), 'enter-probe.go');
            const busy = path.join(workspaceRoot(), 'enter-probe.busy');
            for (const marker of [ready, go, busy]) { fs.rmSync(marker, { force: true }); }
            let blocked = false;
            const selectionEvents: unknown[] = [];
            const selectionListener = vscode.window.onDidChangeTextEditorSelection(event => {
                if (event.textEditor === editor) selectionEvents.push({ kind: event.kind, line: event.selections[0]?.active.line, character: event.selections[0]?.active.character });
            });
            const listener = vscode.workspace.onDidChangeTextDocument(event => {
                if (blocked || event.document !== editor.document || event.contentChanges.length !== 1) { return; }
                const text = event.contentChanges[0].text.trim();
                if (kind === 'comment' ? text !== "'" : text !== '.') { return; }
                blocked = true;
                const busyUntil = Date.now() + 1200;
                fs.writeFileSync(busy, JSON.stringify({ busyUntil }));
                while (Date.now() < busyUntil) { /* test-only pending edit acknowledgement */ }
            });
            const child = spawn(process.execPath, ['-e', probe, String(port), ready, go, busy, kind, move ? '1' : '0', action === 'type' ? '1' : '0'], {
                env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
            });
            let stdout = '', stderr = '';
            child.stdout.on('data', data => { stdout += String(data); });
            child.stderr.on('data', data => { stderr += String(data); });
            const finished = new Promise<number | null>((resolve, reject) => { child.once('error', reject); child.once('close', resolve); });
            void finished.catch(() => {});
            try {
                await until(() => { assert.equal(stderr, '', stderr); return fs.existsSync(ready) || undefined; }, 'renderer Enter probe should become ready', 12000);
                fs.writeFileSync(go, 'go');
                assert.equal(await finished, 0, stderr);
                assert.equal(blocked, true, 'automation edit must pause acknowledgement');
                const result = JSON.parse(stdout.trim());
                console.log('Renderer Enter ownership:', JSON.stringify({ kind, action, selectionEvents, ...result }));
                if (move) {
                    assert.equal(result.movedWhileBusy, true);
                    assert.equal(result.settled.top, result.navigated.top, 'successful Enter acknowledgement must retain native navigation');
                    assert.equal(editor.selection.active.line, openerLine);
                    if (action === 'type') {
                        assert.equal(result.typedWhileBusy, true);
                        assert.ok(editor.document.lineAt(openerLine).text.includes('z'));
                    }
                } else {
                    assert.equal(editor.selection.active.line, openerLine + 1);
                    assert.equal(editor.selection.active.character, editor.document.lineAt(openerLine + 1).text.length);
                }
            } finally {
                listener.dispose(); selectionListener.dispose(); child.kill();
                for (const marker of [ready, go, busy]) { fs.rmSync(marker, { force: true }); }
            }
        });
    }
});
