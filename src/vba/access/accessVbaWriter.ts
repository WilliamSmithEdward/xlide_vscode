import { compress, decompress } from '../ovba';
import { decodeCodePage, encodeCodePage } from '../codePages';
import { readDirRecords, VbaProject } from '../vbaProject';
import { accessVbaCfb } from '../accessDatabase';
import type { Cfb } from '../cfb';
import { decodeRowValues, isAccessFile, type AccessValue } from './accessFormat';
import { AccessPageStore } from './accessPageStore';
import { AccessTable, type AccessRowId } from './accessTableWriter';
import { decodeLongValue, readLongValueFrom } from './accessLongValue';
import { splitRow } from './accessRow';
import type { AccessScalar } from './accessValue';
import {
	readAccessCatalog,
	readAccessStorage,
	type AccessStorageEntry,
} from './accessStorage';
import { buildAccessDesign, parseAccessDesign, type AccessDesign } from './accessDesign';
import { updateTypeInfo } from './accessTypeInfo';
import {
	accessDesignTemplate,
	availablePrototypes,
	withDesignGuid,
} from './accessDesignTemplates';
import type { AccessDesignPrototypes } from './accessDesignEdit';
import { designPrototypes, type AccessDesignKind } from './accessDesignEdit';
import {
	PROP_DATA,
	addToDir,
	addToDirData,
	addToFolderList,
	addToProject,
	addToProjectWm,
	attributeLines,
	dirDataEntries,
	moduleDirBlock,
	moduleOffsetAt,
	newStreamRowName,
	nextFolderName,
	removeFromDir,
	removeFromDirData,
	removeFromFolderList,
	removeFromProject,
	removeFromProjectWm,
	renameDirData,
	renameInDir,
	renameProject,
	renameProjectWm,
	splitModuleSource,
	type AccessModuleKind,
} from './accessVbaStreams';

/**
 * Writing the VBA project of an Access database
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/65).
 *
 * Access keeps no CFB: one `MSysAccessStorage` row per stream, under a tree of
 * folders, and a module costs rows in five of them plus rows in three catalog
 * tables. Every one of those places has to agree - a module listed in one and
 * missing from another is a module Access will show and then refuse to open.
 *
 * Writing takes the SOURCE ROUTE, which is what Access's own `/decompile`
 * does. `_VBA_PROJECT` is the compiled project, and its Version field says
 * which build of VBA compiled it; a version the host does not recognise makes
 * VBA discard the cache and compile from the source in the module streams. So
 * a module's stream is the compressed source alone with its MODULEOFFSET at
 * zero, and none of the compiled tables have to be generated. The cost is that
 * the next open recompiles, and the project's existing source has to compile -
 * a stale cache no longer hides a module that does not.
 *
 * The rules here were measured by pyOpenVBA against Access itself; see its
 * `access/_vba.py` and `docs/access_engine.md`.
 */

export class AccessVbaWriteError extends Error {}

/**
 * Where a new module's unpredictable parts come from. A storage row name is 28
 * random capitals and a module's MODULEEND2 cookie is two random bytes, and
 * every row Access writes carries the time it was written; the defaults are
 * the real ones, and a caller that needs a reproducible file supplies its own.
 */
export interface AccessVbaWriterOptions {
	random?: () => number;
	/** The OLE automation serial to stamp rows with. */
	now?: () => number;
}

/** [MS-OVBA] dir record ids this writer reads or patches. */
const REC_PROJECTCODEPAGE = 0x0003;
const REC_MODULENAME = 0x0019;
const REC_MODULENAME_UNICODE = 0x0047;
const REC_MODULESTREAMNAME = 0x001a;
const REC_MODULESTREAMNAME_UNICODE = 0x0032;
const REC_MODULEOFFSET = 0x0031;
const REC_MODULETYPE_PROCEDURAL = 0x0021;
const REC_MODULETYPE_CLASS = 0x0022;

/** `_VBA_PROJECT` opens `cc 61 <u16 version>`. */
const CACHE_SIGNATURE = Buffer.from([0xcc, 0x61]);
const CACHE_VERSION_AT = 2;
/**
 * Any version the host does not know will do. This one is just below the
 * version Access 2016 writes, so it can never collide with a real build.
 */
const STALE_VERSION = 0x0099;

/** A storage row's Type: 1 is a folder, 2 a stream. */
const TYPE_FOLDER = 1;
const TYPE_STREAM = 2;
const STORAGE_TABLE = 'MSysAccessStorage';
/** The container's listing stream, whose name opens with a 0x03 byte. */
const DIR_DATA = 'DirData';

/** `MSysObjects.Type` for a module, and the type a container row carries. */
const OBJECT_MODULE = -32761;
const OBJECT_CONTAINER = 3;
/** `MSysObjects.Type` for a form and for a report. */
const OBJECT_TYPES: Readonly<Record<AccessDesignKind, number>> = {
	form: -32768, report: -32764,
};
/** The navigation pane's own type for each. */
const NAV_DESIGN_TYPES: Readonly<Record<AccessDesignKind, number>> = {
	form: 32768, report: 32772,
};
/** The storage container and the catalog container each kind lives under. */
const DESIGN_CONTAINERS: Readonly<Record<AccessDesignKind, string>> = {
	form: 'Forms', report: 'Reports',
};
/**
 * A design's ids step by one, where a module's step by four: the object ids
 * Access hands out are per kind, and a design takes the next one down.
 */
const DESIGN_ID_STEP = 1;
/** The navigation pane's own type for a module, and the group it files under. */
const NAV_MODULE_TYPE = 32775;
const NAV_MODULE_GROUP = 8;
/**
 * Access hands out object ids four at a time. Taking max + 1 lands inside the
 * range another object holds, and `AllModules(i).Name` then fails.
 */
const OBJECT_ID_STEP = 4;
const MAX_MODULE_NAME = 64;

/** One form or report, with the design it is described by. */
export interface AccessDesignEntry {
	name: string;
	kind: AccessDesignKind;
	/** The numbered folder under `Forms` or `Reports`. */
	ordinal: string;
	design: AccessDesign;
}

/** One row of `MSysAccessStorage`, decoded. */
interface StorageRow {
	id: number;
	parentId: number;
	name: string;
	type: number;
	rowId: AccessRowId;
	/** The stream's bytes, read from its long value. */
	bytes?: Buffer;
}

/**
 * A database's VBA project, open for writing. Every change goes straight into
 * the page store; `toBuffer` gives the file back.
 */
export class AccessVbaWriter {
	private readonly store: AccessPageStore;
	private readonly storage: AccessTable;
	private readonly random: () => number;
	private readonly now: () => number;

	constructor(data: Buffer, options: AccessVbaWriterOptions = {}) {
		if (!isAccessFile(data)) {
			throw new AccessVbaWriteError('Not a Jet 4 or ACE database.');
		}
		this.store = new AccessPageStore(data);
		this.random = options.random ?? Math.random;
		this.now = options.now ?? accessNow;
		this.storage = this.table(STORAGE_TABLE, data);
	}

	// -- reading ---------------------------------------------------------------

	/** Every storage row as it stands now, read straight from the pages. */
	private rows(): StorageRow[] {
		const definition = this.storage.definition;
		const lv = definition.columns.find((column) => column.name === 'Lv');
		const out: StorageRow[] = [];
		for (const row of this.storage.rows()) {
			const values = decodeRowValues(row.bytes, definition);
			const id = numberOf(values?.get('Id'));
			const name = values?.get('Name');
			if (id === undefined || typeof name !== 'string') {
				continue;
			}
			const entry: StorageRow = {
				id,
				parentId: numberOf(values!.get('ParentId')) ?? -1,
				name,
				type: numberOf(values!.get('Type')) ?? 0,
				rowId: row.id,
			};
			const raw = lv && splitRow(definition, row.bytes).values.get(lv.number);
			if (raw) {
				entry.bytes = readLongValueFrom(this.store, decodeLongValue(raw));
			}
			out.push(entry);
		}
		return out;
	}

	/** The `Modules`, `VBAProject` and module-stream folder ids, walked by name. */
	private folderIds(rows: StorageRow[]): { modules: number; project: number; streams: number } {
		const child = (parent: number, name: string): number => {
			const found = rows.find((row) => row.parentId === parent && row.name === name
				&& row.type === TYPE_FOLDER && row.id !== parent);
			if (!found) {
				throw new AccessVbaWriteError(`${STORAGE_TABLE} has no ${name} folder.`);
			}
			return found.id;
		};
		const root = rows.find((row) => row.name === `${STORAGE_TABLE}_ROOT`);
		if (!root) {
			throw new AccessVbaWriteError('This database has no VBA project.');
		}
		const project = child(child(root.id, 'VBA'), 'VBAProject');
		return { modules: child(root.id, 'Modules'), project, streams: child(project, 'VBA') };
	}

	private dirRow(rows: StorageRow[]): StorageRow {
		const found = rows.find((row) => row.name === 'dir' && row.bytes);
		if (!found) {
			throw new AccessVbaWriteError('This database holds no dir stream.');
		}
		return found;
	}

	/** The decompressed dir stream, and the project's declared code page. */
	private dir(rows: StorageRow[]): { raw: Buffer; codePage: number } {
		const raw = decompress(this.dirRow(rows).bytes!, 'dir');
		return { raw, codePage: projectCodePage(raw) };
	}

	/** The modules the dir stream lists, in the order it lists them. */
	moduleNames(): string[] {
		const rows = this.rows();
		const { raw, codePage } = this.dir(rows);
		return moduleBlocks(raw, codePage).map((block) => block.name);
	}

	/** A module's text as the stream holds it: attribute block and body. */
	moduleText(name: string): string {
		const rows = this.rows();
		const { raw, codePage } = this.dir(rows);
		const { block, stream } = this.moduleStream(rows, raw, codePage, name);
		return decodeCodePage(
			decompress(stream.bytes!.subarray(block.offset), block.streamName), codePage,
		);
	}

	// -- writing ---------------------------------------------------------------

	/**
	 * Replace a module's text. `text` is the whole module as the VBE shows it,
	 * attribute block included; text arriving without one keeps the block the
	 * module already has, so a body pasted in from elsewhere does not silently
	 * strip a class's `VB_PredeclaredId`.
	 */
	setModuleText(name: string, text: string): void {
		const rows = this.rows();
		const { raw, codePage } = this.dir(rows);
		const { block, stream } = this.moduleStream(rows, raw, codePage, name);
		const existing = decodeCodePage(
			decompress(stream.bytes!.subarray(block.offset), block.streamName), codePage,
		);
		const lines = withoutClassPreamble(text.replace(/\r\n?|\n/g, '\r\n').split('\r\n'));
		if (!lines[0]?.startsWith('Attribute ')) {
			lines.unshift(...splitModuleSource(existing).attributes);
		}
		this.writeStream(stream, compress(encodeCodePage(lines.join('\r\n'), codePage)));
		// The stream is now the source alone, so the offset that used to skip
		// the compiled prefix has to go with it.
		const dir = Buffer.from(raw);
		dir.writeUInt32LE(0, block.offsetAt);
		this.writeStream(this.dirRow(rows), compress(dir));
		this.invalidateCache(rows);
	}

	/**
	 * Add a module. It costs a storage folder with its `PropData`, a stream
	 * row, entries in `dir`, `PROJECT`, `PROJECTwm`, the container's
	 * `\x03DirData` and folder list, and a catalog row with its two
	 * navigation-pane rows.
	 */
	addModule(
		name: string,
		text = 'Option Compare Database',
		kind: AccessModuleKind = 'module',
	): void {
		assertModuleName(name);
		let rows = this.rows();
		const { raw: dirRaw, codePage } = this.dir(rows);
		if (moduleBlocks(dirRaw, codePage)
			.some((block) => block.name.toLowerCase() === name.toLowerCase())) {
			throw new AccessVbaWriteError(`A module named ${name} already exists.`);
		}
		const ids = this.folderIds(rows);
		const streamName = newStreamRowName(
			new Set(rows.filter((row) => row.parentId === ids.streams).map((row) => row.name)),
			this.random,
		);
		// Every module carries its own MODULEEND2 word; two sharing one is not
		// something Access writes.
		const cookie = Buffer.from([
			Math.floor(this.random() * 256), Math.floor(this.random() * 256),
		]);
		const folder = nextFolderName('Modules', new Set(
			rows.filter((row) => row.parentId === ids.modules && row.type === TYPE_FOLDER)
				.map((row) => row.name),
		));
		const when = this.now();

		// Ids come from the table's own AutoNumber, not from max + 1: every
		// database Access wrote has the counter equal to its highest id, and
		// leaving it behind makes Access's own next insert collide.
		const folderRow = this.storage.insertNamedRow(new Map<string, AccessScalar>([
			['ParentId', ids.modules], ['Name', folder], ['Type', TYPE_FOLDER],
			['DateCreate', when], ['DateUpdate', when],
		]));
		const folderId = this.idOf(folderRow);
		this.storage.insertNamedRow(new Map<string, AccessScalar>([
			['ParentId', folderId], ['Name', 'PropData'], ['Type', TYPE_STREAM],
			['Lv', PROP_DATA], ['DateCreate', when], ['DateUpdate', when],
		]));
		const source = [...attributeLines(name, kind), ...normalizeBody(text)].join('\r\n');
		this.storage.insertNamedRow(new Map<string, AccessScalar>([
			['ParentId', ids.streams], ['Name', streamName], ['Type', TYPE_STREAM],
			['Lv', compress(encodeCodePage(source, codePage))],
			['DateCreate', when], ['DateUpdate', when],
		]));

		rows = this.rows();
		for (const row of rows) {
			if (!row.bytes || row.bytes.length === 0) {
				continue;
			}
			if (row.name === DIR_DATA && row.parentId === ids.modules) {
				this.writeStream(row, addToDirData(row.bytes, name, folder));
			} else if (row.name === 'PropData' && row.parentId === ids.modules) {
				this.writeStream(row, addToFolderList(row.bytes, folder));
			} else if (row.name === 'PROJECTwm' && row.parentId === ids.project) {
				this.writeStream(row, addToProjectWm(row.bytes, name, codePage));
			} else if (row.name === 'PROJECT' && row.parentId === ids.project) {
				this.writeStream(row, encodeCodePage(
					addToProject(decodeCodePage(row.bytes, codePage), name, kind), codePage,
				));
			}
		}
		this.writeStream(
			this.dirRow(rows),
			compress(addToDir(dirRaw, moduleDirBlock(name, streamName, cookie, kind, codePage))),
		);
		this.addCatalogRows(name, when);
		this.invalidateCache(this.rows());
	}

	/** Rename a module in all eight places its name lives. */
	renameModule(name: string, newName: string): void {
		assertModuleName(newName);
		const from = this.renameModuleStreams(name, newName);
		const ids = this.folderIds(this.rows());
		const listing = this.rows().find(
			(row) => row.name === DIR_DATA && row.parentId === ids.modules && row.bytes?.length,
		);
		if (listing) {
			this.writeStream(listing, renameDirData(listing.bytes!, from, newName));
		}
		this.renameCatalogRows(from, newName);
		this.invalidateCache(this.rows());
	}

	/**
	 * Rename a module wherever the VBA project itself names it: the two `dir`
	 * records, the module's own `Attribute VB_Name`, `PROJECT` and `PROJECTwm`.
	 * Returns the name as the project spelled it, which is what the container's
	 * listing and the catalog have to be looked up by.
	 *
	 * The container's `\x03DirData` is deliberately not touched here. It lists
	 * modules under `Modules` and forms under `Forms`, and a form's code module
	 * is in neither: the project carries `Form_Calculator` while the listing
	 * carries `Calculator`.
	 */
	private renameModuleStreams(name: string, newName: string): string {
		let rows = this.rows();
		const { raw: dirRaw, codePage } = this.dir(rows);
		const blocks = moduleBlocks(dirRaw, codePage);
		const block = blocks.find((entry) => entry.name.toLowerCase() === name.toLowerCase());
		if (!block) {
			throw new AccessVbaWriteError(`The VBA project has no module named ${name}.`);
		}
		if (newName.toLowerCase() !== block.name.toLowerCase()
			&& blocks.some((entry) => entry.name.toLowerCase() === newName.toLowerCase())) {
			throw new AccessVbaWriteError(`A module named ${newName} already exists.`);
		}
		const ids = this.folderIds(rows);
		const stream = this.moduleStream(rows, dirRaw, codePage, block.name).stream;
		const text = decodeCodePage(
			decompress(stream.bytes!.subarray(block.offset), block.streamName), codePage,
		);
		const { attributes, body } = splitModuleSource(text);
		// Rewrite whichever VB_Name line is there rather than matching the old
		// name: the source is ANSI, so a name the code page cannot hold is
		// already the '?'-folded projection in that line and would never match.
		let named = false;
		const renamed = attributes.map((line) => {
			if (named || !/^Attribute VB_Name = "/.test(line)) {
				return line;
			}
			named = true;
			return `Attribute VB_Name = "${newName}"`;
		});
		if (!named) {
			renamed.unshift(`Attribute VB_Name = "${newName}"`);
		}
		this.writeStream(
			stream, compress(encodeCodePage([...renamed, ...body].join('\r\n'), codePage)),
		);
		const dir = renameInDir(dirRaw, block.name, newName, codePage);
		dir.writeUInt32LE(0, moduleOffsetAt(dir, newName, codePage));
		this.writeStream(this.dirRow(rows), compress(dir));

		rows = this.rows();
		for (const row of rows) {
			if (!row.bytes || row.bytes.length === 0 || row.parentId !== ids.project) {
				continue;
			}
			if (row.name === 'PROJECTwm') {
				this.writeStream(row, renameProjectWm(row.bytes, block.name, newName, codePage));
			} else if (row.name === 'PROJECT') {
				const before = decodeCodePage(row.bytes, codePage);
				const after = renameProject(before, block.name, newName);
				if (after !== before) {
					this.writeStream(row, encodeCodePage(after, codePage));
				}
			}
		}
		return block.name;
	}

	/** Remove a module and every structure it occupies. */
	deleteModule(name: string): void {
		this.deleteModuleRows(name, true);
		this.invalidateCache(this.rows());
	}

	/**
	 * Take a module out of the VBA project, and out of the `Modules` container
	 * too when it has a folder there.
	 *
	 * A module behind a form or report has none: it is listed in the project's
	 * `dir` and in `PROJECT` as a `DocClass`, and nowhere under `Modules`. It
	 * goes when its design does, which is what `owned` is false for.
	 */
	private deleteModuleRows(name: string, owned: boolean): string {
		const rows = this.rows();
		const { raw: dirRaw, codePage } = this.dir(rows);
		const blocks = moduleBlocks(dirRaw, codePage);
		const block = blocks.find((entry) => entry.name.toLowerCase() === name.toLowerCase());
		if (!block) {
			throw new AccessVbaWriteError(`The VBA project has no module named ${name}.`);
		}
		const ids = this.folderIds(rows);
		// The container's own listing names each object's folder, which is the
		// link to follow: a module's position among the modules is not it,
		// because the module behind a form or report has no folder at all and
		// still counts in the dir stream.
		const listing = rows.find(
			(row) => row.name === DIR_DATA && row.parentId === ids.modules && row.bytes,
		);
		const named = listing && dirDataEntries(listing.bytes!)
			.find((entry) => entry.name === block.name)?.folder;
		const folder = named === undefined ? undefined : rows.find(
			(row) => row.parentId === ids.modules && row.type === TYPE_FOLDER && row.name === named,
		);
		if (owned && !folder) {
			throw new AccessVbaWriteError(
				`The storage lists no folder for module ${block.name}; a module behind a form or `
				+ 'report has none, and is removed with its design.',
			);
		}
		const doomed = new Set<StorageRow>(folder && owned ? [folder] : []);
		for (const row of rows) {
			if ((folder && owned && row.parentId === folder.id)
				|| (row.parentId === ids.streams && row.name === block.streamName)) {
				doomed.add(row);
			}
		}
		for (const row of rows) {
			if (doomed.has(row) || !row.bytes || row.bytes.length === 0) {
				continue;
			}
			if (row.name === 'dir') {
				this.writeStream(row, compress(removeFromDir(dirRaw, block.name, codePage)));
			} else if (owned && row.name === DIR_DATA && row.parentId === ids.modules) {
				this.writeStream(row, removeFromDirData(row.bytes, block.name));
			} else if (owned && row.name === 'PropData' && row.parentId === ids.modules) {
				this.writeStream(row, removeFromFolderList(row.bytes, folder!.name));
			} else if (row.name === 'PROJECTwm' && row.parentId === ids.project) {
				this.writeStream(row, removeFromProjectWm(row.bytes, block.name, codePage));
			} else if (row.name === 'PROJECT' && row.parentId === ids.project) {
				const before = decodeCodePage(row.bytes, codePage);
				const after = removeFromProject(before, block.name);
				if (after !== before) {
					this.writeStream(row, encodeCodePage(after, codePage));
				}
			}
		}
		// An emptied storage page stays alive and owned, which is what the
		// engine does when a catalog row goes.
		for (const row of doomed) {
			this.storage.deleteRow(row.rowId, false);
		}
		if (owned) {
			this.deleteCatalogRows(block.name);
		}
		return block.name;
	}

	toBuffer(): Buffer {
		return this.store.toBuffer();
	}

	// -- forms and reports -----------------------------------------------------

	/**
	 * Every form and report the database holds. A design lives in a `Blob`
	 * under a numbered folder under `Forms` or `Reports`, and the container's
	 * `\x03DirData` is what names it.
	 */
	designs(): AccessDesignEntry[] {
		const rows = this.rows();
		const out: AccessDesignEntry[] = [];
		for (const [container, kind] of [['Forms', 'form'], ['Reports', 'report']] as const) {
			const folder = rows.find((row) => row.name === container && row.type === TYPE_FOLDER);
			if (!folder) {
				continue;
			}
			const listing = rows.find(
				(row) => row.name === DIR_DATA && row.parentId === folder.id && row.bytes,
			);
			const names = new Map(listing
				? dirDataEntries(listing.bytes!).map((entry) => [entry.folder, entry.name])
				: []);
			for (const ordinal of rows) {
				if (ordinal.parentId !== folder.id || ordinal.type !== TYPE_FOLDER
					|| !/^\d+$/.test(ordinal.name)) {
					continue;
				}
				const children = rows.filter((row) => row.parentId === ordinal.id);
				const blob = children.find((row) => row.name === 'Blob');
				if (!blob?.bytes?.length) {
					continue;
				}
				out.push({
					name: names.get(ordinal.name) ?? ordinal.name,
					kind,
					ordinal: ordinal.name,
					design: parseAccessDesign(blob.bytes),
				});
			}
		}
		return out;
	}

	/**
	 * Edit one form or report. `renamed` maps a control's old name to its new
	 * one so the `TypeInfo` stream carries the change the way Access does -
	 * without it a renamed control keeps its old member and the code behind
	 * the form stops compiling against the new name.
	 */
	editDesign(
		name: string,
		edit: (design: AccessDesign) => AccessDesign,
		renamed: ReadonlyMap<string, string> = new Map(),
	): void {
		const rows = this.rows();
		// A design's text is in the project's code page, as everything else is.
		const { codePage } = this.dir(rows);
		const found = this.designs().find(
			(entry) => entry.name.toLowerCase() === name.toLowerCase(),
		);
		if (!found) {
			throw new AccessVbaWriteError(`This database has no form or report named ${name}.`);
		}
		const container = found.kind === 'form' ? 'Forms' : 'Reports';
		const folder = rows.find((row) => row.name === container && row.type === TYPE_FOLDER)!;
		const ordinal = rows.find(
			(row) => row.parentId === folder.id && row.name === found.ordinal,
		)!;
		const children = rows.filter((row) => row.parentId === ordinal.id);
		const blob = children.find((row) => row.name === 'Blob')!;
		const design = edit(found.design);
		this.writeStream(blob, buildAccessDesign(design));
		const typeInfo = children.find((row) => row.name === 'TypeInfo');
		if (typeInfo?.bytes?.length) {
			this.writeStream(
				typeInfo, updateTypeInfo(found.kind, design, typeInfo.bytes, codePage, renamed),
			);
		}
	}

	/**
	 * The prototypes a design can draw on when a control is added: the ones it
	 * carries, and the captured ones for every type it does not.
	 */
	designPrototypesFor(name: string): AccessDesignPrototypes {
		const found = this.designs().find(
			(entry) => entry.name.toLowerCase() === name.toLowerCase(),
		);
		if (!found) {
			throw new AccessVbaWriteError(`This database has no form or report named ${name}.`);
		}
		return availablePrototypes(found.kind, designPrototypes(found.design));
	}

	/**
	 * Add an empty form or report. The design itself comes from a template
	 * captured from Access, with a GUID of its own patched into the blob and
	 * into the catalog row that repeats it.
	 */
	addDesign(name: string, kind: AccessDesignKind = 'form'): void {
		assertModuleName(name);
		const container = DESIGN_CONTAINERS[kind];
		let rows = this.rows();
		if (this.designs().some((entry) => entry.name.toLowerCase() === name.toLowerCase())) {
			throw new AccessVbaWriteError(`A ${kind} named ${name} already exists.`);
		}
		const folderRow = rows.find(
			(row) => row.name === container && row.type === TYPE_FOLDER,
		);
		if (!folderRow) {
			throw new AccessVbaWriteError(`${STORAGE_TABLE} has no ${container} folder.`);
		}
		const folder = nextFolderName(container, new Set(
			rows.filter((row) => row.parentId === folderRow.id && row.type === TYPE_FOLDER)
				.map((row) => row.name),
		));
		const when = this.now();
		const guid = Buffer.from(
			Array.from({ length: 16 }, () => Math.floor(this.random() * 256)),
		);
		const template = accessDesignTemplate(kind);
		const { blob, catalogProperties } = withDesignGuid(kind, guid);

		const ordinal = this.storage.insertNamedRow(new Map<string, AccessScalar>([
			['ParentId', folderRow.id], ['Name', folder], ['Type', TYPE_FOLDER],
			['DateCreate', when], ['DateUpdate', when],
		]));
		const ordinalId = this.idOf(ordinal);
		// `BlobDelta` is written empty: Access keeps one beside every design.
		for (const [stream, payload] of [
			['Blob', blob], ['TypeInfo', template.typeInfo],
			['BlobDelta', undefined], ['PropData', template.propData],
		] as Array<[string, Buffer | undefined]>) {
			const values = new Map<string, AccessScalar>([
				['ParentId', ordinalId], ['Name', stream], ['Type', TYPE_STREAM],
				['DateCreate', when], ['DateUpdate', when],
			]);
			if (payload) {
				values.set('Lv', payload);
			}
			this.storage.insertNamedRow(values);
		}

		rows = this.rows();
		for (const [stream, add] of [
			[DIR_DATA, (payload: Buffer): Buffer => addToDirData(payload, name, folder)],
			['PropData', (payload: Buffer): Buffer => addToFolderList(payload, folder)],
		] as Array<[string, (payload: Buffer) => Buffer]>) {
			const row = rows.find(
				(entry) => entry.parentId === folderRow.id && entry.name === stream,
			);
			if (row) {
				this.writeStream(row, add(row.bytes ?? Buffer.alloc(4)));
			} else {
				this.storage.insertNamedRow(new Map<string, AccessScalar>([
					['ParentId', folderRow.id], ['Name', stream], ['Type', TYPE_STREAM],
					['Lv', add(Buffer.alloc(4))], ['DateCreate', when], ['DateUpdate', when],
				]));
			}
		}
		this.addDesignCatalogRows(name, kind, when, catalogProperties);
	}

	/**
	 * Rename a form or report in the four places its name lives: the
	 * container's listing, the catalog row, the navigation pane's row, and the
	 * code module behind it, which Access binds by name as `Form_<name>` or
	 * `Report_<name>`. A design Access has never opened a code window for has
	 * no module, and then there are three.
	 */
	renameDesign(name: string, newName: string): void {
		assertModuleName(newName);
		const found = this.designs().find(
			(entry) => entry.name.toLowerCase() === name.toLowerCase(),
		);
		if (!found) {
			throw new AccessVbaWriteError(`This database has no form or report named ${name}.`);
		}
		if (newName.toLowerCase() !== found.name.toLowerCase()
			&& this.designs().some((entry) => entry.name.toLowerCase() === newName.toLowerCase())) {
			throw new AccessVbaWriteError(`A ${found.kind} named ${newName} already exists.`);
		}
		const container = DESIGN_CONTAINERS[found.kind];
		const module = accessDesignModuleName(found.kind, found.name);
		if (this.moduleNames().some((entry) => entry.toLowerCase() === module.toLowerCase())) {
			this.renameModuleStreams(module, accessDesignModuleName(found.kind, newName));
		}
		const rows = this.rows();
		const folderRow = rows.find(
			(row) => row.name === container && row.type === TYPE_FOLDER,
		)!;
		const listing = rows.find(
			(row) => row.parentId === folderRow.id && row.name === DIR_DATA && row.bytes?.length,
		);
		if (listing) {
			this.writeStream(listing, renameDirData(listing.bytes!, found.name, newName));
		}
		this.renameDesignCatalogRows(found.name, newName, found.kind);
		this.invalidateCache(this.rows());
	}

	/** Remove a form or report and every structure it occupies. */
	deleteDesign(name: string): void {
		const found = this.designs().find(
			(entry) => entry.name.toLowerCase() === name.toLowerCase(),
		);
		if (!found) {
			throw new AccessVbaWriteError(`This database has no form or report named ${name}.`);
		}
		const container = DESIGN_CONTAINERS[found.kind];
		const rows = this.rows();
		const folderRow = rows.find(
			(row) => row.name === container && row.type === TYPE_FOLDER,
		)!;
		const ordinal = rows.find(
			(row) => row.parentId === folderRow.id && row.name === found.ordinal,
		)!;
		for (const [stream, drop] of [
			[DIR_DATA, (payload: Buffer): Buffer => removeFromDirData(payload, found.name)],
			['PropData', (payload: Buffer): Buffer => removeFromFolderList(payload, ordinal.name)],
		] as Array<[string, (payload: Buffer) => Buffer]>) {
			const row = rows.find(
				(entry) => entry.parentId === folderRow.id && entry.name === stream && entry.bytes,
			);
			if (row) {
				this.writeStream(row, drop(row.bytes!));
			}
		}
		for (const row of rows.filter((entry) => entry.parentId === ordinal.id)) {
			this.storage.deleteRow(row.rowId, false);
		}
		this.storage.deleteRow(ordinal.rowId, false);
		// The module behind it goes too, or the project is left naming a
		// DocClass whose design is gone, which Access reads as corrupt.
		const module = accessDesignModuleName(found.kind, found.name);
		if (this.moduleNames().some((entry) => entry.toLowerCase() === module.toLowerCase())) {
			this.deleteModuleRows(module, false);
		}
		this.deleteDesignCatalogRows(found.name, found.kind);
		this.invalidateCache(this.rows());
	}

	private addDesignCatalogRows(
		name: string,
		kind: AccessDesignKind,
		when: number,
		properties: Buffer,
	): void {
		const objects = this.table('MSysObjects');
		const catalog = objects.rows()
			.map((row) => decodeRowValues(row.bytes, objects.definition))
			.filter((values): values is Map<string, AccessValue> => values !== undefined);
		const container = catalog.find(
			(values) => values.get('Name') === DESIGN_CONTAINERS[kind]
				&& numberOf(values.get('Type')) === OBJECT_CONTAINER,
		);
		if (!container) {
			throw new AccessVbaWriteError(
				`The catalog has no ${DESIGN_CONTAINERS[kind]} container.`,
			);
		}
		const owner = catalog.find((values) => Buffer.isBuffer(values.get('Owner')))?.get('Owner');
		const negative = catalog
			.map((values) => numberOf(values.get('Id')) ?? 0)
			.filter((id) => id < 0);
		const objectId = (negative.length > 0 ? Math.max(...negative) : -(2 ** 31)) + DESIGN_ID_STEP;
		objects.insertNamedRow(new Map<string, AccessScalar>([
			['Id', objectId], ['ParentId', numberOf(container.get('Id'))!], ['Name', name],
			['Type', OBJECT_TYPES[kind]], ['Flags', 0],
			['Owner', Buffer.isBuffer(owner) ? owner : null],
			['LvProp', properties],
			['DateCreate', when], ['DateUpdate', when],
		]));
		this.optionalTable('MSysNavPaneObjectIDs')?.insertNamedRow(new Map<string, AccessScalar>([
			['Id', objectId], ['Name', name], ['Type', NAV_DESIGN_TYPES[kind]],
		]));
	}

	private renameDesignCatalogRows(
		name: string,
		newName: string,
		kind: AccessDesignKind,
	): void {
		const objects = this.table('MSysObjects');
		let objectId: number | undefined;
		for (const row of objects.rows()) {
			const values = decodeRowValues(row.bytes, objects.definition);
			if (numberOf(values?.get('Type')) === OBJECT_TYPES[kind]
				&& values?.get('Name') === name) {
				objectId = numberOf(values.get('Id'));
				objects.updateNamedRow(row.id, new Map<string, AccessScalar>([['Name', newName]]));
				break;
			}
		}
		if (objectId === undefined) {
			return;
		}
		const nav = this.optionalTable('MSysNavPaneObjectIDs');
		for (const row of nav?.rows() ?? []) {
			const values = decodeRowValues(row.bytes, nav!.definition);
			if (numberOf(values?.get('Id')) === objectId) {
				nav!.updateNamedRow(row.id, new Map<string, AccessScalar>([['Name', newName]]));
			}
		}
	}

	private deleteDesignCatalogRows(name: string, kind: AccessDesignKind): void {
		const objects = this.table('MSysObjects');
		let objectId: number | undefined;
		for (const row of objects.rows()) {
			const values = decodeRowValues(row.bytes, objects.definition);
			if (numberOf(values?.get('Type')) === OBJECT_TYPES[kind]
				&& values?.get('Name') === name) {
				objectId = numberOf(values.get('Id'));
				objects.deleteRow(row.id, false);
				break;
			}
		}
		if (objectId === undefined) {
			return;
		}
		const nav = this.optionalTable('MSysNavPaneObjectIDs');
		for (const row of nav?.rows() ?? []) {
			const values = decodeRowValues(row.bytes, nav!.definition);
			if (numberOf(values?.get('Id')) === objectId) {
				nav!.deleteRow(row.id, false);
			}
		}
		const groups = this.optionalTable('MSysNavPaneGroupToObjects');
		for (const row of groups?.rows() ?? []) {
			const values = decodeRowValues(row.bytes, groups!.definition);
			if (numberOf(values?.get('ObjectID')) === objectId) {
				groups!.deleteRow(row.id, false);
			}
		}
	}

	// -- the catalog -----------------------------------------------------------

	private table(name: string, data?: Buffer): AccessTable {
		const table = this.optionalTable(name, data);
		if (!table) {
			throw new AccessVbaWriteError(`This database has no ${name} table.`);
		}
		return table;
	}

	/**
	 * A table that need not be there. Access creates the navigation-pane
	 * tables the first time it shows the pane, so a database it has only ever
	 * run has none, and a module that does not appear in a pane the database
	 * does not have is not a problem to write around.
	 */
	private optionalTable(name: string, data?: Buffer): AccessTable | undefined {
		const catalog = readAccessCatalog(data ?? this.store.toBuffer())
			.find((entry) => entry.name.toLowerCase() === name.toLowerCase());
		return catalog ? new AccessTable(this.store, catalog.definitionPage) : undefined;
	}

	private addCatalogRows(name: string, when: number): void {
		const objects = this.table('MSysObjects');
		const catalog = objects.rows()
			.map((row) => decodeRowValues(row.bytes, objects.definition))
			.filter((values): values is Map<string, AccessValue> => values !== undefined);
		const container = catalog.find((values) => values.get('Name') === 'Modules'
			&& numberOf(values.get('Type')) === OBJECT_CONTAINER);
		if (!container) {
			throw new AccessVbaWriteError('The catalog has no Modules container.');
		}
		const owner = catalog.find((values) => numberOf(values.get('Type')) === OBJECT_MODULE
			&& Buffer.isBuffer(values.get('Owner')))?.get('Owner');
		const negative = catalog
			.map((values) => numberOf(values.get('Id')) ?? 0)
			.filter((id) => id < 0);
		const objectId = (negative.length > 0 ? Math.max(...negative) : -(2 ** 31)) + OBJECT_ID_STEP;
		objects.insertNamedRow(new Map<string, AccessScalar>([
			['Id', objectId], ['ParentId', numberOf(container.get('Id'))!], ['Name', name],
			['Type', OBJECT_MODULE], ['Flags', 0],
			['Owner', Buffer.isBuffer(owner) ? owner : null],
			['DateCreate', when], ['DateUpdate', when],
		]));
		this.optionalTable('MSysNavPaneObjectIDs')?.insertNamedRow(new Map<string, AccessScalar>([
			['Id', objectId], ['Name', name], ['Type', NAV_MODULE_TYPE],
		]));
		const groups = this.optionalTable('MSysNavPaneGroupToObjects');
		if (!groups) {
			return;
		}
		const positions = groups.rows()
			.map((row) => decodeRowValues(row.bytes, groups.definition))
			.filter((values) => numberOf(values?.get('GroupID')) === NAV_MODULE_GROUP)
			.map((values) => numberOf(values?.get('Position')) ?? -1);
		groups.insertNamedRow(new Map<string, AccessScalar>([
			['GroupID', NAV_MODULE_GROUP], ['ObjectID', objectId], ['Flags', 0], ['Icon', 0],
			['Position', (positions.length > 0 ? Math.max(...positions) : -1) + 1],
		]));
	}

	private renameCatalogRows(name: string, newName: string): void {
		const objects = this.table('MSysObjects');
		for (const row of objects.rows()) {
			const values = decodeRowValues(row.bytes, objects.definition);
			if (numberOf(values?.get('Type')) === OBJECT_MODULE && values?.get('Name') === name) {
				objects.updateNamedRow(row.id, new Map<string, AccessScalar>([['Name', newName]]));
				break;
			}
		}
		const nav = this.optionalTable('MSysNavPaneObjectIDs');
		for (const row of nav?.rows() ?? []) {
			const values = decodeRowValues(row.bytes, nav!.definition);
			if (values?.get('Name') === name) {
				nav!.updateNamedRow(row.id, new Map<string, AccessScalar>([['Name', newName]]));
				break;
			}
		}
	}

	private deleteCatalogRows(name: string): void {
		const objects = this.table('MSysObjects');
		let objectId: number | undefined;
		for (const row of objects.rows()) {
			const values = decodeRowValues(row.bytes, objects.definition);
			if (numberOf(values?.get('Type')) === OBJECT_MODULE && values?.get('Name') === name) {
				objectId = numberOf(values.get('Id'));
				objects.deleteRow(row.id, false);
				break;
			}
		}
		if (objectId === undefined) {
			return;
		}
		const nav = this.optionalTable('MSysNavPaneObjectIDs');
		for (const row of nav?.rows() ?? []) {
			const values = decodeRowValues(row.bytes, nav!.definition);
			if (numberOf(values?.get('Id')) === objectId) {
				nav!.deleteRow(row.id, false);
			}
		}
		const groups = this.optionalTable('MSysNavPaneGroupToObjects');
		for (const row of groups?.rows() ?? []) {
			const values = decodeRowValues(row.bytes, groups!.definition);
			if (numberOf(values?.get('ObjectID')) === objectId) {
				groups!.deleteRow(row.id, false);
			}
		}
	}

	/**
	 * Mark the compiled project stale so VBA rebuilds it from source, and drop
	 * the `__SRP_` rows, which Access runs in preference to the canonical
	 * p-code.
	 */
	private invalidateCache(rows: StorageRow[]): void {
		const cache = rows.find((row) => row.name === '_VBA_PROJECT' && row.bytes?.length);
		if (!cache) {
			throw new AccessVbaWriteError('This database has no _VBA_PROJECT stream.');
		}
		if (!cache.bytes!.subarray(0, CACHE_SIGNATURE.length).equals(CACHE_SIGNATURE)) {
			throw new AccessVbaWriteError('_VBA_PROJECT does not start with its signature.');
		}
		const stale = Buffer.from(cache.bytes!);
		stale.writeUInt16LE(STALE_VERSION, CACHE_VERSION_AT);
		this.writeStream(cache, stale);
		// In row order: deleting a row shifts the rows below it up over the
		// hole, so the order the rows go in decides the bytes left behind on
		// the pages their values vacated.
		const doomed = rows
			.filter((row) => row.name.startsWith('__SRP_'))
			.sort((a, b) => a.rowId.page - b.rowId.page || a.rowId.slot - b.rowId.slot);
		for (const row of doomed) {
			this.storage.deleteRow(row.rowId, false);
		}
	}

	// -- helpers ---------------------------------------------------------------

	private idOf(rowId: AccessRowId): number {
		const bytes = this.storage.fetchRow(rowId);
		const id = numberOf(bytes && decodeRowValues(bytes, this.storage.definition)?.get('Id'));
		if (id === undefined) {
			throw new AccessVbaWriteError('The row just written has no Id.');
		}
		return id;
	}

	private moduleStream(
		rows: StorageRow[], dir: Buffer, codePage: number, name: string,
	): { block: ModuleBlock; stream: StorageRow } {
		const block = moduleBlocks(dir, codePage)
			.find((entry) => entry.name.toLowerCase() === name.toLowerCase());
		if (!block) {
			throw new AccessVbaWriteError(`The VBA project has no module named ${name}.`);
		}
		const stream = rows.find((row) => row.name === block.streamName && row.bytes);
		if (!stream) {
			throw new AccessVbaWriteError(`The storage has no stream row for module ${name}.`);
		}
		return { block, stream };
	}

	private writeStream(row: StorageRow, bytes: Buffer): void {
		this.storage.setLongValue(row.rowId, 'Lv', bytes);
		row.bytes = bytes;
	}
}

/**
 * The module Access binds a design's code to. It exists only once the code
 * window has been opened, but the name is fixed either way.
 */
export function accessDesignModuleName(kind: AccessDesignKind, name: string): string {
	return `${kind === 'form' ? 'Form' : 'Report'}_${name}`;
}

/**
 * The forms and reports a database holds, by name and kind. Cheaper than
 * `AccessVbaWriter.designs()`, which parses every design blob; a listing only
 * needs what the containers' own listings say is there.
 */
export function readAccessDesignNames(
	data: Buffer,
): Array<{ name: string; kind: AccessDesignKind }> {
	let roots;
	try {
		roots = readAccessStorage(data);
	} catch {
		return [];
	}
	if (!roots) {
		return [];
	}
	const out: Array<{ name: string; kind: AccessDesignKind }> = [];
	const walk = (nodes: readonly AccessStorageEntry[]): void => {
		for (const node of nodes) {
			const kind = node.name === 'Forms' ? 'form' : node.name === 'Reports' ? 'report' : undefined;
			if (kind) {
				const listing = node.children.find(
					(child) => child.name.endsWith(DIR_DATA) && child.bytes,
				);
				const ordinals = new Set(node.children
					.filter((child) => /^\d+$/.test(child.name))
					.map((child) => child.name));
				for (const entry of listing ? dirDataEntries(listing.bytes!) : []) {
					// A listing can outlive the folder it names; the folder is
					// what actually holds the design.
					if (ordinals.has(entry.folder)) {
						out.push({ name: entry.name, kind });
					}
				}
				continue;
			}
			walk(node.children);
		}
	};
	walk(roots);
	return out;
}

/** Replace a module's text and give the database back. */
export function writeAccessModuleSource(data: Buffer, name: string, text: string): Buffer {
	const writer = new AccessVbaWriter(data);
	writer.setModuleText(name, text);
	return writer.toBuffer();
}

/**
 * Put an edited project back into the database it came from.
 *
 * `project` is the synthetic CFB `accessVbaCfb` handed out, after the editor
 * changed it. Only what changed is written, so a save that changed nothing
 * returns the file byte for byte.
 */
export function applyAccessVbaProject(data: Buffer, project: Cfb): Buffer {
	const before = VbaProject.parse(accessVbaCfb(data));
	const after = VbaProject.parse(project);
	const had = new Map(before.modules.map((module) => [module.name.toLowerCase(), module]));
	const has = new Map(after.modules.map((module) => [module.name.toLowerCase(), module]));
	const added = [...has.values()].filter((module) => !had.has(module.name.toLowerCase()));
	const removed = [...had.values()].filter((module) => !has.has(module.name.toLowerCase()));
	const changed = [...has.values()].filter((module) => {
		const original = had.get(module.name.toLowerCase());
		return original !== undefined && original.source !== module.source;
	});
	if (added.length === 0 && removed.length === 0 && changed.length === 0) {
		return data;
	}
	const writer = new AccessVbaWriter(data);
	// A rename arrives as one added and one removed module carrying the same
	// body, since the CFB the editor works on has no rename of its own to
	// report; pairing them keeps the module's catalog row rather than
	// destroying and rebuilding it.
	for (const module of [...added]) {
		const match = removed.find(
			(gone) => stripAttributes(gone.source) === stripAttributes(module.source),
		);
		if (!match) {
			continue;
		}
		writer.renameModule(match.name, module.name);
		added.splice(added.indexOf(module), 1);
		removed.splice(removed.indexOf(match), 1);
	}
	for (const module of removed) {
		writer.deleteModule(module.name);
	}
	for (const module of added) {
		writer.addModule(module.name, module.source, module.kind === 'standard' ? 'module' : 'class');
	}
	for (const module of changed) {
		writer.setModuleText(module.name, module.source);
	}
	return writer.toBuffer();
}

function stripAttributes(source: string): string {
	return splitModuleSource(source.replace(/\r\n?|\n/g, '\r\n')).body.join('\r\n');
}

interface ModuleBlock {
	name: string;
	streamName: string;
	kind: AccessModuleKind;
	/** MODULEOFFSET: where the compressed source starts in the stream. */
	offset: number;
	/** Where that offset's four bytes sit in the dir stream. */
	offsetAt: number;
}

function moduleBlocks(dir: Buffer, codePage: number): ModuleBlock[] {
	const out: ModuleBlock[] = [];
	let block: Partial<ModuleBlock> & { nameUnicode?: string; streamNameUnicode?: string } = {};
	for (const record of readDirRecords(dir)) {
		const payload = dir.subarray(record.dataStart, record.dataEnd);
		switch (record.id) {
			case REC_MODULENAME:
				block = { name: decodeCodePage(payload, codePage) };
				break;
			case REC_MODULENAME_UNICODE:
				block.nameUnicode = payload.toString('utf16le');
				break;
			case REC_MODULESTREAMNAME:
				block.streamName = decodeCodePage(payload, codePage);
				break;
			case REC_MODULESTREAMNAME_UNICODE:
				block.streamNameUnicode = payload.toString('utf16le');
				break;
			case REC_MODULEOFFSET:
				block.offset = payload.length >= 4 ? payload.readUInt32LE(0) : 0;
				block.offsetAt = record.dataStart;
				break;
			case REC_MODULETYPE_PROCEDURAL:
			case REC_MODULETYPE_CLASS:
				if (block.name !== undefined && block.streamName !== undefined
					&& block.offsetAt !== undefined) {
					out.push({
						name: trueName(block.name, block.nameUnicode, codePage),
						streamName: trueName(block.streamName, block.streamNameUnicode, codePage),
						kind: record.id === REC_MODULETYPE_CLASS ? 'class' : 'module',
						offset: block.offset ?? 0,
						offsetAt: block.offsetAt,
					});
					block = {};
				}
				break;
			default:
				break;
		}
	}
	return out;
}

/**
 * A name whose characters the project's code page cannot hold is carried in
 * full only by the unicode record; the ANSI record beside it holds the
 * '?'-folded projection. Where the two show exactly that relationship the
 * unicode one is the real name, so lookups speak it. Anything else - a unicode
 * record absent, empty, or disagreeing in some other way - leaves the ANSI name
 * alone, since only the folding relationship proves which is which.
 */
function trueName(ansi: string, unicode: string | undefined, codePage: number): string {
	if (!unicode || unicode === ansi) {
		return ansi;
	}
	return decodeCodePage(encodeCodePage(unicode, codePage), codePage) === ansi ? unicode : ansi;
}

/**
 * The PROJECTCODEPAGE the dir stream declares. [MS-OVBA] 2.3.4.2.1.4 makes the
 * record mandatory, so the fallback is a damaged-file path rather than a real
 * one, and it matches what the reader assumes.
 */
function projectCodePage(dir: Buffer): number {
	for (const record of readDirRecords(dir)) {
		if (record.id === REC_PROJECTCODEPAGE && record.dataEnd - record.dataStart >= 2) {
			return dir.readUInt16LE(record.dataStart);
		}
	}
	return 1252;
}

/**
 * Drop the `VERSION 1.0 CLASS / BEGIN ... END` preamble a CFB class module
 * carries. An Access class module has none - its stream opens straight at the
 * attributes - so text brought in from a workbook or a `.cls` sheds it here.
 */
function withoutClassPreamble(lines: string[]): string[] {
	if (!/^VERSION\s/i.test(lines[0] ?? '')) {
		return lines;
	}
	const end = lines.findIndex((line) => line.trim().toUpperCase() === 'END');
	return end < 0 ? lines.slice(1) : lines.slice(end + 1);
}

/** A new module's body: whatever attribute block it carried is replaced. */
function normalizeBody(text: string): string[] {
	return splitModuleSource(
		withoutClassPreamble(text.replace(/\r\n?|\n/g, '\r\n').split('\r\n')).join('\r\n'),
	).body;
}

function assertModuleName(name: string): void {
	if (!name || name.length > MAX_MODULE_NAME) {
		throw new AccessVbaWriteError(`A module name is 1 to ${MAX_MODULE_NAME} characters.`);
	}
}

/** The stamp Access puts on a row it writes: now, to the second. */
function accessNow(): number {
	const now = new Date();
	now.setMilliseconds(0);
	return now.getTime() / 86400000 + 25569;
}

function numberOf(value: AccessValue | undefined): number | undefined {
	return typeof value === 'number' ? value : undefined;
}

export type { AccessModuleKind };
