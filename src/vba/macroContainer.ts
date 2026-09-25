// The macro container seam: from file bytes to the CFB holding the VBA
// project, and back, for every container Office puts one in.
//
//   - OOXML packages (.xlsm/.xlsb/.xlam, .docm/.dotm, .pptm/.potm/.ppsm):
//     a zip with the project at <host>/vbaProject.bin.
//   - Legacy compound files (.doc, .xls): the whole file is the CFB and the
//     project's VBA storage sits under Macros / _VBA_PROJECT_CUR - the
//     project parser's storage-agnostic lookups find it as-is.
//   - Legacy PowerPoint (.ppt): the project is an embedded, zlib-compressed
//     CFB inside an ExOleObjStg record of the 'PowerPoint Document' stream
//     (pptContainer.ts owns the persist machinery).
//   - Access (.accdb/.mdb): Jet/ACE pages, reassembled by accessDatabase.ts.
//
// Everything is decided from CONTENT, never the file extension: a renamed
// file classifies as what it is. Writes cover every container whose write
// path is mechanically sound: OOXML packages splice the vbaProject part
// back into the zip, legacy .doc/.xls re-serialize the compound file, and
// .ppt rebuilds its embedded record with the persist offsets shifted.
// Access is the one host that runs the compiled project rather than the
// source, so a source write alone would change nothing: accessVbaWriter
// also marks the compiled cache stale, and Access recompiles on the next
// open.

import { accessVbaCfb, isAccessDatabase } from './accessDatabase';
import { Cfb } from './cfb';
import { pptVbaCfb, pptWriteVbaStorage } from './pptContainer';
import {
	accessDesignModuleName,
	applyAccessVbaProject,
	readAccessDesignNames,
} from './access/accessVbaWriter';
import {
	accessDesignMembers,
	parseAccessDesign,
	type AccessDesign,
	type AccessDesignMember,
} from './access/accessDesign';
import { typeInfoListedNames } from './access/accessTypeInfo';
import { NoVbaProjectError } from './noVbaProject';
import { XlsxWorkbook } from './xlsx';

export class MacroContainerError extends Error {}

// Re-exported so the seam that hands out containers is also the one import
// site for the state a container can be in.
export { NoVbaProjectError } from './noVbaProject';

export type MacroContainerKind = 'excel' | 'word' | 'powerpoint' | 'access';

export interface MacroContainer {
	kind: MacroContainerKind;
	/** False only where writing cannot take effect; every container today can. */
	writable: boolean;
	/** Noun phrase for messages: "a legacy Word document (.doc)". */
	description: string;
	/** The OOXML package, when the container is one (any host). */
	xlsx?: XlsxWorkbook;
	/** The compound file itself, when the document is one (.xls, .doc): its own streams, VBA or not. */
	cfb?: Cfb;
	/** The CFB holding the VBA project; parsed once and cached. Raises
	 * {@link NoVbaProjectError} when the file has no VBA project in it yet. */
	vbaCfb(): Cfb;
	/** The whole container file's bytes with the (mutated) VBA project CFB
	 * spliced back in. Throws for read-only containers. */
	toFileBytes(cfb: Cfb): Buffer;
	/**
	 * Forms and reports the host keeps outside the VBA project. Only Access
	 * has any: its designs are objects in the database, and the module behind
	 * one is optional, so they cannot be found by walking the project alone.
	 */
	designs?(): AccessContainerDesign[];
}

/** One design a container carries, and the module Access names for its code. */
export interface AccessContainerDesign {
	name: string;
	kind: 'form' | 'report';
	/** `Form_<name>` or `Report_<name>`, whether or not that module exists. */
	moduleName: string;
	/**
	 * The design's named sections and controls, which are members of its
	 * class. Parsed on the first ask; undefined when the design could not be
	 * read, which means "not known", never "none". Given the project's code
	 * page, a control Access left out of the design's member list - one whose
	 * name that list's page cannot hold - is left out here too.
	 */
	members(projectCodePage?: number): AccessDesignMember[] | undefined;
}

const ZIP_MAGIC = Buffer.from('PK', 'latin1');
const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

export function openMacroContainer(data: Buffer): MacroContainer {
	if (data.subarray(0, 2).equals(ZIP_MAGIC)) {
		return openOoxmlContainer(data);
	}
	if (data.subarray(0, 8).equals(CFB_MAGIC)) {
		return openLegacyCfbContainer(data);
	}
	if (isAccessDatabase(data)) {
		return {
			kind: 'access',
			writable: true,
			description: 'an Access database',
			vbaCfb: cached(() => accessVbaCfb(data)),
			// Access runs the compiled project, not the source, so a source
			// write only takes effect once the compiled cache is marked stale;
			// the writer does that, and Access recompiles on the next open.
			toFileBytes: (cfb: Cfb): Buffer => applyAccessVbaProject(data, cfb),
			designs: cachedDesigns(data),
		};
	}
	throw new MacroContainerError(
		'Not a macro-enabled Office file: expected an OOXML package, a legacy Office compound file, or an Access database.',
	);
}

function openOoxmlContainer(data: Buffer): MacroContainer {
	const xlsx = XlsxWorkbook.fromBuffer(data);
	const host = xlsx.packageHost() ?? 'excel';
	const descriptions: Record<'excel' | 'word' | 'powerpoint', string> = {
		excel: 'an Excel workbook',
		word: 'a Word macro-enabled document',
		powerpoint: 'a PowerPoint macro-enabled presentation',
	};
	return {
		kind: host,
		writable: true,
		description: descriptions[host],
		xlsx,
		vbaCfb: cached(() => {
			if (!xlsx.hasVbaProject()) {
				throw new NoVbaProjectError(descriptions[host]);
			}
			return Cfb.fromBytes(xlsx.readVbaProject());
		}),
		toFileBytes: (cfb: Cfb): Buffer => {
			xlsx.writeVbaProject(cfb.toBytes());
			return xlsx.toBytes();
		},
	};
}

function openLegacyCfbContainer(data: Buffer): MacroContainer {
	const outer = Cfb.fromBytes(data);
	if (outer.hasStream('WordDocument')) {
		return wholeCfbContainer(outer, 'word', 'a legacy Word document (.doc)');
	}
	if (outer.hasStream('Workbook') || outer.hasStream('Book')) {
		return wholeCfbContainer(outer, 'excel', 'a legacy Excel workbook (.xls)');
	}
	if (outer.hasStream('PowerPoint Document')) {
		return {
			kind: 'powerpoint',
			writable: true,
			description: 'a legacy PowerPoint presentation (.ppt)',
			vbaCfb: cached(() => pptVbaCfb(outer)),
			toFileBytes: (cfb: Cfb): Buffer => pptWriteVbaStorage(outer, cfb.toBytes()).toBytes(),
		};
	}
	if (hasVbaDirStream(outer)) {
		// A bare VBA project: a legacy PowerPoint add-in (.ppa) saves as
		// exactly this shape - a VBA storage and PROJECT at the root with no
		// document stream at all - and a stray vbaProject.bin is the same.
		// The analysis host still comes from the file name, so the neutral
		// kind here only shapes refusal wording for surfaces the file lacks.
		return wholeCfbContainer(outer, 'powerpoint', 'a VBA project compound file');
	}
	throw new MacroContainerError(
		'Compound file without a recognizable Office host (no WordDocument, Workbook, or PowerPoint Document stream, and no VBA project at the root).',
	);
}

/**
 * Where a project's streams live in a compound file: under a `VBA` storage
 * (.doc and .xls keep that storage inside Macros / _VBA_PROJECT_CUR, which
 * the storage-agnostic lookup finds) or at the root, which is the shape of a
 * bare vbaProject.bin. Neither means the file has never had any code.
 */
function hasVbaDirStream(cfb: Cfb): boolean {
	return cfb.hasStreamInStorage('VBA', 'dir') || cfb.hasStream('dir');
}

/** .doc / .xls: the file IS the CFB, so writing is re-serializing it. */
function wholeCfbContainer(outer: Cfb, kind: MacroContainerKind, description: string): MacroContainer {
	return {
		kind,
		writable: true,
		description,
		cfb: outer,
		vbaCfb: (): Cfb => {
			// A legacy document with no macros carries no VBA storage at all;
			// parsing it would answer "not a valid VBA project", which reads
			// as damage rather than as the empty file it is.
			if (!hasVbaDirStream(outer)) {
				throw new NoVbaProjectError(description);
			}
			return outer;
		},
		toFileBytes: (cfb: Cfb): Buffer => cfb.toBytes(),
	};
}

/**
 * The database's forms and reports, read once. A design with no module behind
 * it is still a design, so the list comes from the storage rather than from
 * the project's module list.
 */
function cachedDesigns(data: Buffer): () => AccessContainerDesign[] {
	let value: AccessContainerDesign[] | undefined;
	return (): AccessContainerDesign[] => {
		value ??= readAccessDesignNames(data).map(({ name, kind, blob, typeInfo }) => ({
			name,
			kind,
			moduleName: accessDesignModuleName(kind, name),
			members: designMembers(blob, kind, typeInfo),
		}));
		return value;
	};
}

/** A design's members, parsed once for a code page and only when asked for. */
function designMembers(
	blob: Buffer | undefined,
	kind: 'form' | 'report',
	typeInfo: Buffer | undefined,
): (projectCodePage?: number) => AccessDesignMember[] | undefined {
	const parsed = new Map<number | undefined, AccessDesignMember[] | undefined>();
	return (projectCodePage) => {
		if (!parsed.has(projectCodePage)) {
			let value: AccessDesignMember[] | undefined;
			try {
				const design = blob ? parseAccessDesign(blob) : undefined;
				value = design && accessDesignMembers(design, kind, listedNames(design, typeInfo, projectCodePage));
			} catch {
				// A design this cannot read still has a module; its members
				// are simply not known.
				value = undefined;
			}
			parsed.set(projectCodePage, value);
		}
		return parsed.get(projectCodePage);
	};
}

/** The names a design's member list holds, or undefined where it cannot say. */
function listedNames(
	design: AccessDesign,
	typeInfo: Buffer | undefined,
	projectCodePage: number | undefined,
): Set<string> | undefined {
	if (!typeInfo || projectCodePage === undefined) {
		return undefined;
	}
	try {
		return typeInfoListedNames(typeInfo, design, projectCodePage);
	} catch {
		// A member list this cannot read says nothing about what it holds.
		return undefined;
	}
}

function cached(build: () => Cfb): () => Cfb {
	let value: Cfb | undefined;
	let builder: (() => Cfb) | undefined = build;
	return (): Cfb => {
		if (value === undefined && builder !== undefined) {
			value = builder();
			// Release the builder so its closure (for Access, the entire
			// database file buffer - Cfb.addStream already copied the stream
			// bytes out of it) does not stay reachable for the cache's life.
			builder = undefined;
		}
		if (value === undefined) {
			throw new MacroContainerError('VBA project could not be built for this container.');
		}
		return value;
	};
}
