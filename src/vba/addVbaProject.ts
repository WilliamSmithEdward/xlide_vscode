// Adds a VBA project to an Office Open XML file that has none: a workbook,
// document or presentation saved in a macro-enabled format before its first
// macro was written, which carries no vbaProject.bin part at all.
//
// Excel, Word and PowerPoint add a project themselves the first time code is
// written in the VBE, and what XLIDE writes follows what they wrote when
// asked to, measured on Office 16.0 (build 20326) against files each had
// saved without VBA:
//
//   - Excel makes a document module for the workbook and for every worksheet
//     and chart sheet, in tab order after ThisWorkbook, and writes each code
//     name into the package: `codeName` on `workbookPr`, and on each sheet's
//     `sheetPr`, which becomes the sheet's first element - or, in a binary
//     workbook, into the empty name of each property record (below). A sheet
//     with no code name gets one from its kind and its position among the
//     worksheets and chart sheets (Chart1, Sheet2, Sheet3 for a chart sheet
//     in front of two worksheets); dialog sheets and Excel 4 macro sheets get
//     neither a code name nor a module, and do not count.
//   - Word makes ThisDocument.
//   - PowerPoint makes a project with no modules in it.
//   - Each declares the part as `<Default Extension="bin">` with the VBA
//     project content type, and relates it to the main part.
//
// The project itself - its references, code page, name, protection lines -
// comes from the blank template XLIDE creates new files from, the one for
// the file's own format, with its starter module taken out.

import { Cfb } from './cfb';
import { VbaProject } from './vbaProject';
import { ZipArchive } from './zip';

export class AddVbaProjectError extends Error {}

const VBA_PROJECT_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';
const VBA_PROJECT_RELATIONSHIP = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const CONTENT_TYPES = '[Content_Types].xml';

/**
 * The main part's content type for each macro-enabled extension, read from
 * files Office saved in each format (the templates and fixtures). A package
 * whose main part says otherwise - a renamed .xlsx - cannot be opened by its
 * application under that extension, VBA or not, so it is corrected.
 */
const MACRO_ENABLED_MAIN: Record<string, { host: OoxmlHost; contentType: string }> = {
    xlsm: { host: 'excel', contentType: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' },
    xltm: { host: 'excel', contentType: 'application/vnd.ms-excel.template.macroEnabled.main+xml' },
    xlam: { host: 'excel', contentType: 'application/vnd.ms-excel.addin.macroEnabled.main+xml' },
    docm: { host: 'word', contentType: 'application/vnd.ms-word.document.macroEnabled.main+xml' },
    dotm: { host: 'word', contentType: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml' },
    pptm: { host: 'powerpoint', contentType: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml' },
    potm: { host: 'powerpoint', contentType: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml' },
    ppsm: { host: 'powerpoint', contentType: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml' },
    ppam: { host: 'powerpoint', contentType: 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml' },
};

/**
 * VB_Base of a chart sheet's document module: the one attribute in which it
 * differs from a worksheet's, as Excel wrote it into a workbook with a chart
 * sheet that it added VBA to.
 */
const CHART_SHEET_BASE = '0{00020821-0000-0000-C000-000000000046}';

type OoxmlHost = 'excel' | 'word' | 'powerpoint';

/** Where a package keeps its main part, and what relates to it. */
interface MainPart {
    host: OoxmlHost;
    /** e.g. `xl/workbook.xml`. */
    path: string;
    /** e.g. `xl/`. */
    folder: string;
    /** e.g. `xl/_rels/workbook.xml.rels`. */
    rels: string;
}

interface Relationship {
    id: string;
    type: string;
    target: string;
}

/** A sheet that gets a document module, in tab order. */
interface CodedSheet {
    path: string;
    kind: 'worksheet' | 'chartsheet';
    codeName?: string;
}

export interface AddedVbaProject {
    /** The file with its new project. */
    bytes: Buffer;
    /** The modules the project starts with, in project order. */
    modules: string[];
}

/**
 * Adds a VBA project to the package in `data` - a file with extension
 * `extension`, lowercase, no dot - using the project in `template`, the
 * blank file of the same format.
 */
export function addVbaProjectToPackage(data: Buffer, extension: string, template: Buffer): AddedVbaProject {
    const zip = ZipArchive.read(data);
    const main = mainPart(zip);
    if (zip.names().some((name) => /(^|\/)vbaProject\.bin$/i.test(name))) {
        throw new AddVbaProjectError('The file already has a VBA project.');
    }

    const templateZip = ZipArchive.read(template);
    const templatePart = templateZip.names().find((name) => /(^|\/)vbaProject\.bin$/i.test(name));
    if (!templatePart) {
        throw new AddVbaProjectError('The template for this format has no VBA project to start from.');
    }
    const cfb = Cfb.fromBytes(templateZip.read(templatePart));
    const project = VbaProject.parse(cfb);

    if (main.host === 'excel') {
        startExcelProject(zip, main, project);
    } else {
        // Word keeps ThisDocument; PowerPoint has no document module. The
        // starter modules the templates carry go either way.
        for (const module of [...project.modules]) {
            if (!(main.host === 'word' && isDocumentModule(module.source))) {
                project.deleteModule(module.name);
            }
        }
    }
    project.save(cfb);

    const partPath = `${main.folder}vbaProject.bin`;
    zip.write(partPath, cfb.toBytes());
    declareVbaProjectPart(zip, partPath);
    relateVbaProjectPart(zip, main);
    ensureMacroEnabledMain(zip, main, extension);
    return { bytes: zip.toBytes(), modules: project.modules.map((module) => module.name) };
}

/** The package's main part, from its root relationship. */
function mainPart(zip: ZipArchive): MainPart {
    const root = zip.has('_rels/.rels') ? relationships(zip.read('_rels/.rels').toString('utf8')) : [];
    const office = root.find((relationship) => relationship.type.endsWith('/officeDocument'));
    if (!office) {
        throw new AddVbaProjectError('The file is not an Office document: its package names no main part.');
    }
    const path = office.target.replace(/^\//, '');
    const slash = path.lastIndexOf('/');
    const folder = slash >= 0 ? path.slice(0, slash + 1) : '';
    const name = path.slice(slash + 1);
    const host: OoxmlHost | undefined = /^workbook\.(xml|bin)$/i.test(name) ? 'excel'
        : /^document\.xml$/i.test(name) ? 'word'
            : /^presentation\.xml$/i.test(name) ? 'powerpoint'
                : undefined;
    if (!host || !zip.has(path)) {
        throw new AddVbaProjectError(`The file's main part (${path}) is not a workbook, document or presentation.`);
    }
    return { host, path, folder, rels: `${folder}_rels/${name}.rels` };
}

function relationships(xml: string): Relationship[] {
    const out: Relationship[] = [];
    for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
        const attrs = attributes(match[1]);
        if (attrs.Id && attrs.Type && attrs.Target) {
            out.push({ id: attrs.Id, type: attrs.Type, target: attrs.Target });
        }
    }
    return out;
}

function attributes(text: string): Record<string, string> {
    const out: Record<string, string> = {};
    for (const match of text.matchAll(/([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g)) {
        out[match[1]] = match[3] ?? match[4] ?? '';
    }
    return out;
}

function isDocumentModule(source: string): boolean {
    return /^\s*Attribute\s+VB_Base\s*=/im.test(source);
}

/** The attribute header of a module, with its name and, when given, its base replaced. */
function documentHeader(prototype: string, name: string, base?: string): string {
    const header = prototype.split(/\r?\n/).filter((line) => /^\s*Attribute\s+VB_/i.test(line));
    return header
        .map((line) => line
            .replace(/^(\s*Attribute\s+VB_Name\s*=\s*")[^"]*(")/i, `$1${name}$2`)
            .replace(/^(\s*Attribute\s+VB_Base\s*=\s*")[^"]*(")/i, (whole, open: string, close: string) => (base ? `${open}${base}${close}` : whole)))
        .join('\r\n') + '\r\n';
}

/**
 * Excel's part: the workbook's module and one per worksheet and chart sheet,
 * with the code names written into the package - as XML attributes in a
 * workbook, or into the property records of a binary workbook.
 */
function startExcelProject(zip: ZipArchive, main: MainPart, project: VbaProject): void {
    const workbookModule = project.modules.find((module) => /VB_Base\s*=\s*"0\{00020819-/i.test(module.source));
    const sheetModule = project.modules.find((module) => /VB_Base\s*=\s*"0\{00020820-/i.test(module.source));
    if (!workbookModule || !sheetModule) {
        throw new AddVbaProjectError('The Excel template has no workbook and worksheet modules to start from.');
    }
    const sheetPrototype = sheetModule.source;
    const book = main.path.endsWith('.bin') ? binaryWorkbook(zip, main) : xmlWorkbook(zip, main);
    const workbookCodeName = book.workbookCodeName || 'ThisWorkbook';
    assignCodeNames(book.sheets, workbookCodeName);

    for (const module of [...project.modules]) {
        if (module !== workbookModule) {
            project.deleteModule(module.name);
        }
    }
    if (workbookModule.name !== workbookCodeName) {
        project.renameModule(workbookModule.name, workbookCodeName);
    }
    for (const sheet of book.sheets) {
        const codeName = sheet.codeName!;
        const header = documentHeader(sheetPrototype, codeName, sheet.kind === 'chartsheet' ? CHART_SHEET_BASE : undefined);
        project.addModule(codeName, header, 'other', { projectKeyword: 'Document' });
    }
    book.writeCodeNames(workbookCodeName);
}

/** A workbook's sheets that take code names, and how to write the names into it. */
interface ExcelBook {
    workbookCodeName?: string;
    sheets: CodedSheet[];
    writeCodeNames(workbookCodeName: string): void;
}

/** A sheet's relationship from the workbook, as the part it names and its kind. */
function sheetPart(main: MainPart, relationship: Relationship | undefined): { path: string; kind: CodedSheet['kind'] } | undefined {
    const kind = relationship?.type.endsWith('/worksheet') ? 'worksheet'
        : relationship?.type.endsWith('/chartsheet') ? 'chartsheet'
            : undefined;
    // Dialog sheets and Excel 4 macro sheets: no code name, no module.
    if (!relationship || !kind) {
        return undefined;
    }
    const path = relationship.target.startsWith('/') ? relationship.target.slice(1) : `${main.folder}${relationship.target}`;
    return { path, kind };
}

function mainRelationships(zip: ZipArchive, main: MainPart): Map<string, Relationship> {
    return new Map(relationships(zip.has(main.rels) ? zip.read(main.rels).toString('utf8') : '')
        .map((relationship) => [relationship.id, relationship]));
}

/** An Office Open XML workbook: code names are attributes of workbookPr and sheetPr. */
function xmlWorkbook(zip: ZipArchive, main: MainPart): ExcelBook {
    const workbookXml = zip.read(main.path).toString('utf8');
    const rels = mainRelationships(zip, main);
    const sheetsBlock = /<sheets\b[^>]*>([\s\S]*?)<\/sheets>/.exec(workbookXml)?.[1] ?? '';
    const sheets: CodedSheet[] = [];
    for (const match of sheetsBlock.matchAll(/<sheet\b([^>]*)\/?>/g)) {
        const attrs = attributes(match[1]);
        const part = sheetPart(main, rels.get(attrs['r:id'] ?? attrs.id ?? ''));
        if (!part || !zip.has(part.path)) {
            continue;
        }
        const sheetPr = /<sheetPr\b([^>]*)\/?>/.exec(zip.read(part.path).toString('utf8'));
        const codeName = sheetPr ? attributes(sheetPr[1]).codeName : undefined;
        sheets.push({ ...part, ...(codeName ? { codeName } : {}) });
    }
    return {
        workbookCodeName: attributes(/<workbookPr\b([^>]*)\/?>/.exec(workbookXml)?.[1] ?? '').codeName,
        sheets,
        writeCodeNames(workbookCodeName: string): void {
            zip.write(main.path, Buffer.from(withWorkbookCodeName(workbookXml, workbookCodeName), 'utf8'));
            for (const sheet of sheets) {
                const xml = zip.read(sheet.path).toString('utf8');
                zip.write(sheet.path, Buffer.from(withSheetCodeName(xml, sheet.codeName!), 'utf8'));
            }
        },
    };
}

/**
 * Names each sheet without a code name the way Excel does: its kind and its
 * position among the sheets that have one. A name already taken - a code
 * name kept from before, or the workbook's - moves on to the next number.
 */
function assignCodeNames(sheets: CodedSheet[], workbookCodeName: string): void {
    const taken = new Set([workbookCodeName.toLowerCase(), ...sheets.flatMap((sheet) => (sheet.codeName ? [sheet.codeName.toLowerCase()] : []))]);
    sheets.forEach((sheet, index) => {
        if (sheet.codeName) {
            return;
        }
        const prefix = sheet.kind === 'chartsheet' ? 'Chart' : 'Sheet';
        let number = index + 1;
        while (taken.has(`${prefix}${number}`.toLowerCase())) {
            number += 1;
        }
        sheet.codeName = `${prefix}${number}`;
        taken.add(sheet.codeName.toLowerCase());
    });
}

/** `workbookPr` with a code name, added where the workbook has none. */
function withWorkbookCodeName(xml: string, codeName: string): string {
    const existing = /<workbookPr\b([^>]*?)(\/?)>/.exec(xml);
    if (existing) {
        return attributes(existing[1]).codeName
            ? xml
            : xml.replace(existing[0], `<workbookPr codeName="${codeName}"${existing[1]}${existing[2]}>`);
    }
    // workbookPr follows fileVersion and fileSharing and comes before
    // everything else Excel writes at the top of the workbook.
    const next = /<(?:mc:AlternateContent|xr:revisionPtr|workbookProtection|bookViews|sheets)\b/.exec(xml);
    return next ? `${xml.slice(0, next.index)}<workbookPr codeName="${codeName}"/>${xml.slice(next.index)}` : xml;
}

/** `sheetPr` with a code name, made the sheet's first element where it has none. */
function withSheetCodeName(xml: string, codeName: string): string {
    const existing = /<sheetPr\b([^>]*?)(\/?)>/.exec(xml);
    if (existing) {
        return attributes(existing[1]).codeName
            ? xml
            : xml.replace(existing[0], `<sheetPr codeName="${codeName}"${existing[1]}${existing[2]}>`);
    }
    const root = /<(?:worksheet|chartsheet)\b[^>]*>/.exec(xml);
    if (!root) {
        return xml;
    }
    const at = root.index + root[0].length;
    return `${xml.slice(0, at)}<sheetPr codeName="${codeName}"/>${xml.slice(at)}`;
}

// ------------------------------------------------------------ binary workbooks
//
// An .xlsb keeps the same code names in records ([MS-XLSB] 2.1.4): each is a
// type and a size, both variable-length integers of 7 bits a byte, low byte
// first, with the high bit set when another byte follows, then the payload.
// Excel's own "add VBA" to a binary workbook changed exactly these records,
// filling each empty code name, and one byte offset that moved because of it.

const BRT_WB_PROP = 153;
const BRT_BUNDLE_SH = 156;
const BRT_WS_PROP = 147;
const BRT_CS_PROP = 651;
const BRT_INDEX_ROW_BLOCK = 40;
const BINARY_INDEX_RELATIONSHIP = 'http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex';

/**
 * Where each property record's code name starts: after the flags and theme
 * version of BrtWbProp (2.4.866), the flags, tab colour and scroll anchors of
 * BrtWsProp (2.4.875), and the flags and tab colour of BrtCsProp (2.4.344).
 */
const CODE_NAME_AT: Record<number, number> = { [BRT_WB_PROP]: 8, [BRT_WS_PROP]: 19, [BRT_CS_PROP]: 10 };

interface BiffRecord {
    type: number;
    start: number;
    payloadStart: number;
    end: number;
}

function readVarint(data: Buffer, at: number, maxBytes: number): { value: number; next: number } {
    let value = 0;
    for (let i = 0; i < maxBytes; i++) {
        if (at + i >= data.length) {
            break;
        }
        const byte = data[at + i];
        value += (byte & 0x7f) * 2 ** (7 * i);
        if (!(byte & 0x80)) {
            return { value, next: at + i + 1 };
        }
    }
    throw new AddVbaProjectError('The binary workbook has a record XLIDE cannot read.');
}

function varint(value: number): number[] {
    const bytes: number[] = [];
    let rest = value;
    do {
        const low = rest % 128;
        rest = Math.floor(rest / 128);
        bytes.push(rest > 0 ? low | 0x80 : low);
    } while (rest > 0);
    return bytes;
}

/** The records of a binary part, which must end exactly where the part does. */
function biffRecords(data: Buffer): BiffRecord[] {
    const out: BiffRecord[] = [];
    let at = 0;
    while (at < data.length) {
        const type = readVarint(data, at, 2);
        const size = readVarint(data, type.next, 4);
        const end = size.next + size.value;
        if (end > data.length) {
            throw new AddVbaProjectError('The binary workbook has a record that runs past the end of its part.');
        }
        out.push({ type: type.value, start: at, payloadStart: size.next, end });
        at = end;
    }
    return out;
}

function encodeRecord(type: number, payload: Buffer): Buffer {
    return Buffer.concat([Buffer.from([...varint(type), ...varint(payload.length)]), payload]);
}

/** An XLWideString: a 4-byte character count, then UTF-16. */
function readWideString(data: Buffer, at: number, end: number): { text: string; end: number } | undefined {
    if (at + 4 > end) {
        return undefined;
    }
    const count = data.readUInt32LE(at);
    const stop = at + 4 + count * 2;
    if (count === 0xffffffff || stop > end) {
        return undefined;
    }
    return { text: data.toString('utf16le', at + 4, stop), end: stop };
}

function wideString(text: string): Buffer {
    const count = Buffer.alloc(4);
    count.writeUInt32LE(text.length, 0);
    return Buffer.concat([count, Buffer.from(text, 'utf16le')]);
}

/** The first record of a type, and the code name it carries. */
function codeNameRecord(data: Buffer, type: number): { record: BiffRecord; codeName: string } | undefined {
    const record = biffRecords(data).find((candidate) => candidate.type === type);
    if (!record) {
        return undefined;
    }
    const name = readWideString(data, record.payloadStart + CODE_NAME_AT[type], record.end);
    return { record, codeName: name?.text ?? '' };
}

/**
 * The part with its property record's code name set, and how far everything
 * after the record moved. What follows the name in the payload, if anything,
 * is kept as it was.
 */
function withRecordCodeName(data: Buffer, type: number, codeName: string): { data: Buffer; moved: { from: number; by: number } } {
    const found = codeNameRecord(data, type);
    if (!found) {
        throw new AddVbaProjectError('The binary workbook has a sheet without the record that holds its code name.');
    }
    const { record } = found;
    const at = record.payloadStart + CODE_NAME_AT[type];
    const oldName = readWideString(data, at, record.end);
    const after = oldName ? oldName.end : Math.min(at + 4, record.end);
    const payload = Buffer.concat([data.subarray(record.payloadStart, at), wideString(codeName), data.subarray(after, record.end)]);
    const rewritten = encodeRecord(type, payload);
    return {
        data: Buffer.concat([data.subarray(0, record.start), rewritten, data.subarray(record.end)]),
        moved: { from: record.end, by: rewritten.length - (record.end - record.start) },
    };
}

/**
 * Moves the worksheet offsets a binary index holds past `moved.from` by
 * `moved.by`: each BrtIndexRowBlock (2.4.699) locates its first cell by
 * ibBaseOffset, a 64-bit byte index into the worksheet part after a 4-byte
 * row mask; the offsets after it are relative to it and stay as they are.
 */
function withShiftedIndex(data: Buffer, moved: { from: number; by: number }): Buffer {
    const out = Buffer.from(data);
    for (const record of biffRecords(data)) {
        if (record.type !== BRT_INDEX_ROW_BLOCK || record.end - record.payloadStart < 12) {
            continue;
        }
        const at = record.payloadStart + 4;
        const offset = out.readBigUInt64LE(at);
        if (offset >= BigInt(moved.from)) {
            out.writeBigUInt64LE(offset + BigInt(moved.by), at);
        }
    }
    return out;
}

/** The binary index a worksheet part relates to, if it has one. */
function binaryIndexOf(zip: ZipArchive, sheetPath: string): string | undefined {
    const slash = sheetPath.lastIndexOf('/');
    const folder = sheetPath.slice(0, slash + 1);
    const rels = `${folder}_rels/${sheetPath.slice(slash + 1)}.rels`;
    if (!zip.has(rels)) {
        return undefined;
    }
    const index = relationships(zip.read(rels).toString('utf8')).find((relationship) => relationship.type === BINARY_INDEX_RELATIONSHIP);
    if (!index) {
        return undefined;
    }
    const path = index.target.startsWith('/') ? index.target.slice(1) : `${folder}${index.target}`;
    return zip.has(path) ? path : undefined;
}

/** A binary workbook: code names are strings in BrtWbProp, BrtWsProp and BrtCsProp. */
function binaryWorkbook(zip: ZipArchive, main: MainPart): ExcelBook {
    const workbook = zip.read(main.path);
    const rels = mainRelationships(zip, main);
    const sheets: CodedSheet[] = [];
    // BrtBundleSh (2.4.316), in tab order: hsState, itabID, then the sheet's
    // relationship id and its name, each a 4-byte count and UTF-16.
    for (const record of biffRecords(workbook).filter((candidate) => candidate.type === BRT_BUNDLE_SH)) {
        const relId = readWideString(workbook, record.payloadStart + 8, record.end);
        const part = sheetPart(main, relId ? rels.get(relId.text) : undefined);
        if (!part || !zip.has(part.path)) {
            continue;
        }
        const codeName = codeNameRecord(zip.read(part.path), part.kind === 'chartsheet' ? BRT_CS_PROP : BRT_WS_PROP)?.codeName;
        sheets.push({ ...part, ...(codeName ? { codeName } : {}) });
    }
    return {
        workbookCodeName: codeNameRecord(workbook, BRT_WB_PROP)?.codeName || undefined,
        sheets,
        writeCodeNames(workbookCodeName: string): void {
            if (!codeNameRecord(workbook, BRT_WB_PROP)?.codeName) {
                zip.write(main.path, withRecordCodeName(workbook, BRT_WB_PROP, workbookCodeName).data);
            }
            for (const sheet of sheets) {
                const data = zip.read(sheet.path);
                const type = sheet.kind === 'chartsheet' ? BRT_CS_PROP : BRT_WS_PROP;
                if (codeNameRecord(data, type)?.codeName) {
                    continue;
                }
                const edited = withRecordCodeName(data, type, sheet.codeName!);
                zip.write(sheet.path, edited.data);
                const index = binaryIndexOf(zip, sheet.path);
                if (index && edited.moved.by !== 0) {
                    zip.write(index, withShiftedIndex(zip.read(index), edited.moved));
                }
            }
        },
    };
}

/**
 * The content type for the part: `Default Extension="bin"`, as Office writes
 * it, unless the package already gives .bin parts another type, in which
 * case the part gets an override of its own.
 */
function declareVbaProjectPart(zip: ZipArchive, partPath: string): void {
    const xml = zip.read(CONTENT_TYPES).toString('utf8');
    const binDefault = /<Default\b[^>]*Extension="bin"[^>]*\/>/i.exec(xml);
    let edited: string;
    if (binDefault) {
        edited = binDefault[0].includes(VBA_PROJECT_CONTENT_TYPE)
            ? xml
            : xml.replace(/<\/Types>/, `<Override PartName="/${partPath}" ContentType="${VBA_PROJECT_CONTENT_TYPE}"/></Types>`);
    } else {
        edited = xml.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="bin" ContentType="${VBA_PROJECT_CONTENT_TYPE}"/>`);
    }
    zip.write(CONTENT_TYPES, Buffer.from(edited, 'utf8'));
}

/** Relates the part to the main part, under the next free id. */
function relateVbaProjectPart(zip: ZipArchive, main: MainPart): void {
    const xml = zip.has(main.rels)
        ? zip.read(main.rels).toString('utf8')
        : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const used = relationships(xml).map((relationship) => /^rId(\d+)$/.exec(relationship.id)?.[1]).filter((id): id is string => id !== undefined);
    const id = `rId${Math.max(0, ...used.map(Number)) + 1}`;
    const relationship = `<Relationship Id="${id}" Type="${VBA_PROJECT_RELATIONSHIP}" Target="vbaProject.bin"/>`;
    zip.write(main.rels, Buffer.from(xml.replace(/<\/Relationships>/, `${relationship}</Relationships>`), 'utf8'));
}

/** Gives the main part the macro-enabled content type its extension needs. */
function ensureMacroEnabledMain(zip: ZipArchive, main: MainPart, extension: string): void {
    const wanted = MACRO_ENABLED_MAIN[extension];
    if (!wanted || wanted.host !== main.host) {
        return;
    }
    const xml = zip.read(CONTENT_TYPES).toString('utf8');
    const override = new RegExp(`<Override\\b[^>]*PartName="/${main.path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`, 'i').exec(xml);
    if (!override || override[0].includes(`ContentType="${wanted.contentType}"`)) {
        return;
    }
    const fixed = override[0].replace(/ContentType="[^"]*"/, `ContentType="${wanted.contentType}"`);
    zip.write(CONTENT_TYPES, Buffer.from(xml.replace(override[0], fixed), 'utf8'));
}
