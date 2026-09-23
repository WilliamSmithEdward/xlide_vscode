// Shape looks, checked by the applications themselves: each file XLIDE
// formats is opened in its application with alerts on, beside a watcher that
// records and closes any dialog it raises, and the application reports every
// shape's fill, outline, font, rotation, visibility and stacking through its
// own object model. The unit suites compare XLIDE's markup with what each
// application saved; this is the other direction, the application reading
// what XLIDE wrote.
//
// Excel runs in instances the check starts itself. Word and PowerPoint only
// when neither is running, so no instance of theirs is ever touched. Enum
// values are the repo's Office models': msoTrue -1, msoFalse 0, msoLineDash
// 4, msoUnderlineSingleLine 2, wdUnderlineSingle 1, ppAlertsAll 2,
// wdAlertsAll -1.

import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { editShape } from '../../src/vba/projectService';
import type { ShapeEdit } from '../../src/vba/shapes';
import { dialogsSeen, fields, hostInstalled, hostIsFree, psFile, q, scratchCopy, watchLines } from './officeHarness';

/** What an application reported for one shape. */
interface Reported {
    name: string;
    fillVisible?: number;
    fillRGB?: number;
    transparency?: number;
    lineVisible?: number;
    lineRGB?: number;
    weight?: number;
    dash?: number;
    rotation?: number;
    visible?: number;
    z?: number;
    text?: string;
    fontName?: string;
    fontSize?: number;
    bold?: number;
    italic?: number;
    underline?: number;
    fontRGB?: number;
}

/** An Office RGB value: red in the low byte. */
const rgb = (hex: string): number => {
    const n = parseInt(hex.replace('#', ''), 16);
    return ((n >> 16) & 255) + ((n >> 8) & 255) * 256 + (n & 255) * 65536;
};

/**
 * PowerShell reporting a shape as one JSON line. Every read is guarded: a
 * group has no text frame, a form control no DrawingML fill.
 */
const REPORT = [
    'function Report($s, $textFrame) {',
    '  $f = [ordered]@{ name = $s.Name }',
    '  try { $f.fillVisible = [int]$s.Fill.Visible; $f.fillRGB = [int]$s.Fill.ForeColor.RGB; $f.transparency = [math]::Round([double]$s.Fill.Transparency, 2) } catch { }',
    '  try { $f.lineVisible = [int]$s.Line.Visible; $f.lineRGB = [int]$s.Line.ForeColor.RGB; $f.weight = [double]$s.Line.Weight; $f.dash = [int]$s.Line.DashStyle } catch { }',
    '  try { $f.rotation = [double]$s.Rotation } catch { }',
    '  try { $f.visible = [int]$s.Visible } catch { }',
    '  try { $f.z = [int]$s.ZOrderPosition } catch { }',
    '  try {',
    '    if ($textFrame -eq 2) { $t = $s.TextFrame2.TextRange; $ft = $t.Font; $f.underline = [int]$ft.UnderlineStyle; $f.fontRGB = [int]$ft.Fill.ForeColor.RGB }',
    '    else { $t = $s.TextFrame.TextRange; $ft = $t.Font; $f.underline = [int]$ft.Underline; $f.fontRGB = [int]$ft.Color }',
    '    $f.text = ([string]$t.Text).TrimEnd([char]13, [char]10); $f.fontName = $ft.Name; $f.fontSize = [double]$ft.Size; $f.bold = [int]$ft.Bold; $f.italic = [int]$ft.Italic',
    '  } catch { }',
    '  "shape=" + (ConvertTo-Json $f -Compress)',
    '}',
];

function reported(lines: readonly string[]): Map<string, Reported> {
    return new Map(lines.filter((line) => line.startsWith('shape=')).map((line) => {
        const shape = JSON.parse(line.slice('shape='.length)) as Reported;
        return [shape.name, shape];
    }));
}

/** The look XLIDE gave the main shape in each file, as the application should report it. */
const FORMATTED = {
    fillVisible: -1, fillRGB: rgb('#FF0000'), transparency: 0.25,
    lineVisible: -1, lineRGB: rgb('#008000'), weight: 2.5, dash: 4,
    rotation: 30, fontName: 'Arial', fontSize: 14, bold: -1, italic: -1, fontRGB: rgb('#0000FF'),
};

const LOOK: Pick<ShapeEdit, 'fill' | 'line' | 'font' | 'rotation'> = {
    rotation: 30,
    fill: { type: 'solid', color: '#FF0000', transparency: 25 },
    line: { type: 'solid', color: '#008000', weight: 2.5, dash: 'dash' },
    font: { name: 'Arial', size: 14, bold: true, italic: true, underline: true, color: '#0000FF' },
};

describe.runIf(await hostInstalled('excel'))('Shape looks, opened in Excel', () => {
    it('reads every look XLIDE wrote, with nothing to repair', async () => {
        const file = scratchCopy('ShapesFixture.xlsm', 'shape-looks-excel');
        editShape(file, 'Sheet1', { action: 'update', name: 'RunButton', ...LOOK });
        editShape(file, 'Sheet1', { action: 'update', name: 'Oval 3', fill: { type: 'none' }, line: { type: 'none' } });
        editShape(file, 'Sheet1', { action: 'update', name: 'TextBox 4', hidden: true });
        editShape(file, 'Sheet1', { action: 'update', name: 'Pair', rotation: 45 });
        editShape(file, 'Sheet1', { action: 'update', name: 'Straight Connector 5', line: { type: 'solid', color: '#800000', weight: 4 } });
        editShape(file, 'Sheet1', { action: 'update', name: 'Drop Down 3', hidden: true });
        editShape(file, 'Sheet1', {
            action: 'add', type: 'oval', name: 'Added', range: 'K2:L4', text: 'Hi', rotation: 90,
            fill: { type: 'solid', color: '#00FF00' }, line: { type: 'none' }, font: { bold: true },
        });
        const log = path.join(path.dirname(file), 'dialogs.log');
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            ...REPORT,
            '$xl = [Activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))',
            ...watchLines('$xl.Hwnd', log),
            '$xl.DisplayAlerts = $true',
            '$wb = $null',
            `try { $wb = $xl.Workbooks.Open(${q(file)}); "opened=True" } catch { "opened=False"; "error=" + $_ }`,
            'if ($wb) {',
            '  foreach ($s in @($wb.Worksheets.Item("Sheet1").Shapes)) { Report $s 2 }',
            '  "topZ=" + (@($wb.Worksheets.Item("Sheet1").Shapes | ForEach-Object { $_.ZOrderPosition }) | Measure-Object -Maximum).Maximum',
            '  $wb.Close($false)',
            '}',
            '$xl.DisplayAlerts = $false',
            '$wb = $null',
            '$xl.Quit()',
            '$xl = $null',
            '[GC]::Collect()',
            '[GC]::WaitForPendingFinalizers()',
            'Start-Sleep -Milliseconds 800',
        ].join('\r\n'), 'excel-shape-looks.ps1');
        const seen = fields(result.out);
        expect(seen.opened, `${seen.error ?? ''} ${result.err.join(' ')}`).toBe('True');
        expect(dialogsSeen(log)).toEqual([]);
        const shapes = reported(result.out);
        expect(shapes.get('RunButton')).toMatchObject(FORMATTED);
        expect(shapes.get('RunButton')?.underline).toBe(2);
        expect(shapes.get('Oval 3')).toMatchObject({ fillVisible: 0, lineVisible: 0 });
        expect(shapes.get('TextBox 4')?.visible).toBe(0);
        expect(shapes.get('Pair')?.rotation).toBe(45);
        expect(shapes.get('Straight Connector 5')).toMatchObject({ lineRGB: rgb('#800000'), weight: 4 });
        expect(shapes.get('Drop Down 3')?.visible).toBe(0);
        expect(shapes.get('Added')).toMatchObject({ fillRGB: rgb('#00FF00'), lineVisible: 0, rotation: 90, text: 'Hi', bold: -1 });
        // Added last, so on top of every other shape on the sheet.
        expect(shapes.get('Added')?.z).toBe(Number(seen.topZ));
    });
});

describe.runIf(await hostIsFree('powerpoint'))('Shape looks, opened in PowerPoint', () => {
    it('reads every look XLIDE wrote, with nothing to repair', async () => {
        const file = scratchCopy('PowerPointShapesFixture.pptm', 'shape-looks-powerpoint');
        editShape(file, 'Slide 1', { action: 'update', name: 'ClickMe', ...LOOK });
        editShape(file, 'Slide 1', { action: 'update', name: 'Badge', fill: { type: 'none' }, line: { type: 'none' }, rotation: 90, zOrder: 'back' });
        editShape(file, 'Slide 1', { action: 'update', name: 'Caption', hidden: true });
        editShape(file, 'Slide 2', {
            action: 'add', type: 'oval', name: 'Dot', left: 10, top: 20, width: 50, height: 40, text: 'Hi', rotation: 45,
            fill: { type: 'solid', color: '#00FF00' }, font: { bold: true },
        });
        const log = path.join(path.dirname(file), 'dialogs.log');
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            ...REPORT,
            '$pp = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))',
            ...watchLines('$pp.HWND', log),
            '$pp.DisplayAlerts = 2',
            '$pres = $null',
            // Presentations.Open(FileName, ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoTrue)
            `try { $pres = $pp.Presentations.Open(${q(file)}, -1, 0, -1); "opened=True" } catch { "opened=False"; "error=" + $_ }`,
            'if ($pres) {',
            '  foreach ($slide in @($pres.Slides)) { foreach ($s in @($slide.Shapes)) { Report $s 2 } }',
            '  $pres.Close()',
            '}',
            '$pres = $null',
            '$pp.Quit()',
            '$pp = $null',
            '[GC]::Collect()',
            '[GC]::WaitForPendingFinalizers()',
        ].join('\r\n'), 'powerpoint-shape-looks.ps1');
        const seen = fields(result.out);
        expect(seen.opened, `${seen.error ?? ''} ${result.err.join(' ')}`).toBe('True');
        expect(dialogsSeen(log)).toEqual([]);
        const shapes = reported(result.out);
        expect(shapes.get('ClickMe')).toMatchObject({ ...FORMATTED, underline: 2 });
        expect(shapes.get('Badge')).toMatchObject({ fillVisible: 0, lineVisible: 0, rotation: 90, z: 1 });
        expect(shapes.get('Caption')?.visible).toBe(0);
        // A new shape's text is the size the slide gives it, as PowerPoint's own are.
        expect(shapes.get('Dot')).toMatchObject({ fillRGB: rgb('#00FF00'), rotation: 45, text: 'Hi', bold: -1, fontSize: 18 });
    });
});

describe.runIf(await hostIsFree('word'))('Shape looks, opened in Word', () => {
    it('reads every look XLIDE wrote, in the canvas too, with nothing to repair', async () => {
        const file = scratchCopy('WordShapesFixture.docm', 'shape-looks-word');
        editShape(file, 'Document', { action: 'update', name: 'AnchoredBox', ...LOOK });
        editShape(file, 'Document', { action: 'update', name: 'Board', hidden: true });
        editShape(file, 'Document', { action: 'update', name: 'GroupedRect', hidden: true });
        editShape(file, 'Document', { action: 'update', name: 'GroupedOval', fill: { type: 'solid', color: '#FFFF00' }, line: { type: 'none' }, rotation: 20 });
        editShape(file, 'Document', { action: 'update', name: 'InlineOval', text: 'Hi' });
        editShape(file, 'Document', {
            action: 'add', type: 'rectangle', name: 'Added', left: 20, top: 300, width: 80, height: 40, text: 'Yo',
            rotation: 45, fill: { type: 'solid', color: '#00FF00' }, font: { bold: true },
        });
        const log = path.join(path.dirname(file), 'dialogs.log');
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            ...REPORT,
            '$wd = [Activator]::CreateInstance([type]::GetTypeFromProgID("Word.Application"))',
            '$wd.Visible = $true',
            '$blank = $wd.Documents.Add()',
            ...watchLines('$wd.ActiveWindow.Hwnd', log),
            '$blank.Close(0)',
            '$wd.DisplayAlerts = -1',
            '$doc = $null',
            // Documents.Open(FileName, ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False)
            `try { $doc = $wd.Documents.Open(${q(file)}, $false, $true, $false); "opened=True" } catch { "opened=False"; "error=" + $_ }`,
            'if ($doc) {',
            '  foreach ($s in @($doc.Shapes)) { Report $s 1 }',
            '  $pair = $doc.Shapes.Item("Board").CanvasItems.Item("Pair")',
            '  foreach ($s in @($pair.GroupItems)) { Report $s 1 }',
            '  $doc.Close(0)',
            '}',
            '$wd.DisplayAlerts = 0',
            '$doc = $null',
            '$pair = $null',
            '$blank = $null',
            '$wd.Quit(0)',
            '$wd = $null',
            '[GC]::Collect()',
            '[GC]::WaitForPendingFinalizers()',
        ].join('\r\n'), 'word-shape-looks.ps1');
        const seen = fields(result.out);
        expect(seen.opened, `${seen.error ?? ''} ${result.err.join(' ')}`).toBe('True');
        expect(dialogsSeen(log)).toEqual([]);
        const shapes = reported(result.out);
        expect(shapes.get('AnchoredBox')).toMatchObject({ ...FORMATTED, underline: 1 });
        expect(shapes.get('Board')?.visible).toBe(0);
        expect(shapes.get('GroupedRect')?.visible).toBe(0);
        expect(shapes.get('GroupedOval')).toMatchObject({ fillRGB: rgb('#FFFF00'), lineVisible: 0, rotation: 20 });
        expect(shapes.get('InlineOval')?.text).toBe('Hi');
        expect(shapes.get('Added')).toMatchObject({ fillRGB: rgb('#00FF00'), rotation: 45, text: 'Yo', bold: -1 });
        // Word counts the inline shape too; the added one is on top of them all.
        const zs = [...shapes.values()].filter((s) => ['InlineOval', 'AnchoredBox', 'Board', 'Added'].includes(s.name)).map((s) => s.z ?? 0);
        expect(shapes.get('Added')?.z).toBe(Math.max(...zs));
    });
});
