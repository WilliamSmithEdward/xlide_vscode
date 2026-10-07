// Real extension-host tests using the permanent password-protected workbook.
// As in deleteModule.test.ts, stub only the user dialog; all file-system,
// editor, save and language-model tool operations use the activated extension.
import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { ProjectEngine } from '../projectEngine';
import { ProjectExplorer } from '../projectExplorer';
import { encodeModuleUri, encodeFormMarkupUri, XlideFileSystemProvider } from '../xlideFileSystem';
import { updateProjectModuleSyncSettings } from '../projectModuleSyncSettings';
import { openMacroContainer } from '../vba/macroContainer';
import { activate, closeAllEditors, EXTENSION_ID, until, workspaceRoot } from './support';

const PASSWORD = 'Test66';
const SOURCE_MARKER = 'counter = 1';

suite('Protected VBA project integration', () => {
    let dir: string;
    let file: string;
    let before: Buffer;
    let originalInput: typeof vscode.window.showInputBox;
    const prompts: vscode.InputBoxOptions[] = [];
    let answers: Array<string | undefined>;
    suiteSetup(activate);
    setup(() => {
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-password-integration-'));
        file = path.join(dir, 'Protected.xlsm');
        const extension = vscode.extensions.getExtension(EXTENSION_ID)!;
        fs.copyFileSync(path.join(extension.extensionPath, 'tests/fixtures/binaries/PasswordProtectedFixture.xlsm'), file);
        before = fs.readFileSync(file);
        prompts.length = 0; answers = [];
        originalInput = vscode.window.showInputBox;
        vscode.window.showInputBox = async options => {
            assert.equal(options?.title, 'Unlock VBA project');
            assert.equal(options?.password, true, 'password entry must be masked');
            assert.equal(options?.ignoreFocusOut, true);
            assert.ok(!Object.prototype.hasOwnProperty.call(options, 'value'), 'never prefill passwords');
            prompts.push(options!);
            return answers.shift();
        };
    });
    teardown(async () => {
        // Revert any failed editor save before close; never leave dirty fixtures.
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
        await closeAllEditors();
        vscode.window.showInputBox = originalInput;
        fs.rmSync(dir, { recursive: true, force: true });
    });
    const moduleUri = () => encodeModuleUri(file, 'Runner');
    async function invoke(tool: string, input: Record<string, unknown>): Promise<string> {
        const result = await vscode.lm.invokeTool(tool, {
            input: { filePath: file, ...input }, toolInvocationToken: undefined,
        }, new vscode.CancellationTokenSource().token);
        return result.content.map(part => part instanceof vscode.LanguageModelTextPart ? part.value : '').join('');
    }

    for (const command of ['xlide.analyzeProject', 'xlide.importModulesFromFolder', 'xlide.exportModulesToFolder', 'xlide.runVbaTests']) {
        test(command + ' stops at the password gate before opening its workflow', async () => {
            answers = [undefined];
            const chooseFolder = vscode.window.showOpenDialog;
            let folderPrompts = 0;
            vscode.window.showOpenDialog = async () => { folderPrompts++; return undefined; };
            const tabsBefore = vscode.window.tabGroups.all.flatMap(group => group.tabs).length;
            try {
                try { await vscode.commands.executeCommand(command, { kind: 'project', filePath: file, label: 'Protected.xlsm' }); }
                catch (err) { assert.ok(err instanceof vscode.CancellationError || /cancel/i.test(String(err))); }
            } finally { vscode.window.showOpenDialog = chooseFolder; }
            assert.ok(prompts.length >= 1, command + ' must request the existing VBA password');
            assert.equal(folderPrompts, 0, 'unlock must happen before selecting an import/export folder');
            assert.equal(vscode.window.tabGroups.all.flatMap(group => group.tabs).length, tabsBefore, 'no preview, analysis or test panel should open');
            assert.deepEqual(fs.readFileSync(file), before);
        });
    }
    test('cancelling a module read reveals no source and changes no bytes', async () => {
        answers = [undefined];
        await assert.rejects(async () => vscode.workspace.fs.readFile(moduleUri()));
        assert.equal(prompts.length, 1);
        assert.deepEqual(fs.readFileSync(file), before);
    });
    test('the tree shows a closed padlock before prompting and an open padlock after unlock', async () => {
        const treeFile = path.join(workspaceRoot(), `PasswordIcon-${path.basename(dir)}.xlsm`);
        fs.copyFileSync(file, treeFile);
        const extension = vscode.extensions.getExtension(EXTENSION_ID)!;
        const engine = new ProjectEngine({} as vscode.ExtensionContext);
        const explorer = new ProjectExplorer(engine, undefined, undefined, vscode.Uri.file(extension.extensionPath));
        try {
            const node = (await explorer.getChildren()).find(row => row.filePath === treeFile)!;
            assert.ok(node, 'real workspace discovery should find the protected fixture');
            explorer.getTreeItem(node);
            await until(() => node.isPasswordProtected, 'collapsed file should acquire its protection badge');
            const icon = () => explorer.getTreeItem(node).iconPath as { light: vscode.Uri; dark: vscode.Uri };
            assert.ok(icon().light.path.endsWith('file-code-locked.svg'));
            assert.ok(fs.existsSync(icon().light.fsPath) && fs.existsSync(icon().dark.fsPath));
            assert.equal(prompts.length, 0, 'badge metadata must never prompt');
            answers = [undefined];
            assert.equal((await explorer.getChildren(node))[0].kind, 'loadError');
            assert.ok(icon().dark.path.endsWith('file-code-locked.svg'));
            answers = [PASSWORD];
            assert.ok((await explorer.getChildren(node)).some(row => row.moduleName === 'Runner'));
            await until(() => node.isAccessLocked === false, 'successful unlock should update the existing tree row');
            assert.ok(icon().dark.path.endsWith('file-code-unlocked.svg'));
            assert.match(String(explorer.getTreeItem(node).description), /unlocked for this session/);
            await explorer.getChildren(node);
            assert.equal(prompts.length, 2, 'repeated expansion must reuse authorization');
            assert.deepEqual(fs.readFileSync(treeFile), before);
        } finally { explorer.dispose(); engine.dispose(); fs.unlinkSync(treeFile); }
    });
    test('cancelling a provider write leaves the protected file intact', async () => {
        answers = [undefined];
        const engine = new ProjectEngine({} as vscode.ExtensionContext);
        const provider = new XlideFileSystemProvider(engine);
        try {
            await assert.rejects(() => provider.writeFile(moduleUri(), Buffer.from('Option Explicit\r\n'), { create: false, overwrite: true }));
        } finally { provider.dispose(); engine.dispose(); }
        assert.ok(prompts.length >= 1);
        assert.deepEqual(fs.readFileSync(file), before);
    });
    test('incorrect passwords retry, then a correct password opens and saves the editor', async () => {
        answers = ['incorrect', PASSWORD];
        const document = await vscode.workspace.openTextDocument(moduleUri());
        await vscode.window.showTextDocument(document, { preview: false });
        assert.ok(document.getText().includes(SOURCE_MARKER));
        assert.equal(prompts.length, 2);
        assert.ok(prompts[1].prompt?.includes('Incorrect password.'));
        const edit = new vscode.WorkspaceEdit();
        edit.replace(document.uri, new vscode.Range(0, 0, document.lineCount, 0), document.getText().replace(SOURCE_MARKER, 'counter = 2'));
        assert.equal(await vscode.workspace.applyEdit(edit), true);
        assert.equal(await document.save(), true);
        const saved = fs.readFileSync(file);
        assert.notDeepEqual(saved, before);
        const protection = (bytes: Buffer) => openMacroContainer(bytes).vbaCfb().getStream('PROJECT').toString('latin1').split(/\r?\n/).filter(line => /^(CMG|DPB|GC)=/.test(line));
        assert.deepEqual(protection(saved), protection(before), 'save must preserve all protection records');
        assert.ok(Buffer.from(await vscode.workspace.fs.readFile(moduleUri())).toString('utf8').includes('counter = 2'));
        assert.equal(prompts.length, 2, 'authorization should be reused in this session');
    });
    test('the protected UserForm markup and designer path also requires a password', async () => {
        const extension = vscode.extensions.getExtension(EXTENSION_ID)!;
        fs.copyFileSync(path.join(extension.extensionPath, 'tests/fixtures/binaries/PasswordProtectedFormFixture.xlsm'), file);
        before = fs.readFileSync(file);
        const uri = encodeFormMarkupUri(file, 'FrmPicker');
        answers = [undefined];
        await assert.rejects(async () => vscode.workspace.fs.readFile(uri));
        answers = [PASSWORD];
        const document = await vscode.workspace.openTextDocument(uri);
        assert.ok(document.getText().length > 0);
        await vscode.commands.executeCommand('vscode.openWith', uri, 'xlideFormDesigner', vscode.ViewColumn.One);
        assert.ok(vscode.window.tabGroups.all.flatMap(group => group.tabs).some(tab => tab.input instanceof vscode.TabInputCustom && tab.input.uri.toString() === uri.toString()));
        assert.deepEqual(fs.readFileSync(file), before);
    });

    test('one unlock is reused across repeated sidebar actions', async () => {
        const folder = path.join(dir, 'modules');
        fs.mkdirSync(folder);
        fs.writeFileSync(path.join(folder, 'ProtectedProbe.bas'), 'Option Explicit\r\nPublic Function ProtectedValue() As Long\r\n    ProtectedValue = 1298\r\nEnd Function\r\n');
        await updateProjectModuleSyncSettings(file, { folderPath: folder });
        answers = [PASSWORD];
        assert.ok(Buffer.from(await vscode.workspace.fs.readFile(moduleUri())).toString('utf8').includes(SOURCE_MARKER));
        const node = { kind: 'project', filePath: file, label: 'Protected.xlsm' };
        for (let repeat = 0; repeat < 2; repeat++) {
            for (const [command, label] of [
                ['xlide.analyzeProject', 'XLIDE Analysis:'],
                ['xlide.importModulesFromFolder', 'XLIDE Import Preview'],
                ['xlide.exportModulesToFolder', 'XLIDE Export Preview'],
                ['xlide.runVbaTests', 'XLIDE Tests:'],
            ]) {
                const finished = Promise.resolve(vscode.commands.executeCommand(command, node));
                // Sync previews wait until the user closes their panel, so observe
                // its real tab and close it through VS Code before awaiting them.
                const tab = await until(() => vscode.window.tabGroups.all.flatMap(group => group.tabs).find(tab => tab.label.startsWith(label)), command + ' should open its normal workflow');
                assert.equal(prompts.length, 1, command + ' must reuse session authorization');
                await vscode.window.tabGroups.close(tab);
                await finished;
            }
        }
        assert.equal(prompts.length, 1);
        assert.deepEqual(fs.readFileSync(file), before);
    });
    test('agent reads cannot bypass a cancelled password prompt', async () => {
        answers = [undefined];
        let result = '';
        try { result = await invoke('xlide_readModule', { moduleName: 'Runner' }); }
        catch { /* Some VS Code versions propagate cancellation instead of tool text. */ }
        assert.ok(prompts.length >= 1, 'the agent must reach the same password gate');
        assert.ok(!result.includes(SOURCE_MARKER), 'source must not appear in the agent result');
        assert.deepEqual(fs.readFileSync(file), before);
    });
    test('agent writes are blocked until the user supplies the existing password', async () => {
        answers = [undefined];
        try { await invoke('xlide_writeModule', { moduleName: 'Runner', source: 'Option Explicit\r\n' }); }
        catch { /* A cancelled tool may reject. */ }
        assert.ok(prompts.length >= 1);
        assert.deepEqual(fs.readFileSync(file), before);
        answers = [PASSWORD];
        const result = await invoke('xlide_readModule', { moduleName: 'Runner' });
        assert.ok(result.includes(SOURCE_MARKER), result);
        const count = prompts.length;
        assert.ok(Buffer.from(await vscode.workspace.fs.readFile(moduleUri())).toString('utf8').includes(SOURCE_MARKER));
        assert.equal(prompts.length, count, 'agent authorization must be shared with editor access');
    });
});
