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
// Access alone stays read-only, for cause: the engine renders and runs VBA
// from its compiled p-code tables, and the MS-OVBA source blob this reader
// extracts is a passive cache - writing source there would silently change
// nothing in Access itself.

import { accessVbaCfb, isAccessDatabase } from './accessDatabase';
import { Cfb } from './cfb';
import { pptVbaCfb, pptWriteVbaStorage } from './pptContainer';
import { XlsxWorkbook } from './xlsx';

export class MacroContainerError extends Error {}

export type MacroContainerKind = 'excel' | 'word' | 'powerpoint' | 'access';

export interface MacroContainer {
	kind: MacroContainerKind;
	/** False only where writing cannot take effect (Access). */
	writable: boolean;
	/** Noun phrase for messages: "a legacy Word document (.doc)". */
	description: string;
	/** The OOXML package, when the container is one (any host). */
	xlsx?: XlsxWorkbook;
	/** The CFB holding the VBA project; parsed once and cached. Throws when
	 * the file has no VBA project. */
	vbaCfb(): Cfb;
	/** The whole container file's bytes with the (mutated) VBA project CFB
	 * spliced back in. Throws for read-only containers. */
	toFileBytes(cfb: Cfb): Buffer;
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
			writable: false,
			description: 'an Access database (read-only: Access runs VBA from its compiled p-code, so source writes would not take effect)',
			vbaCfb: cached(() => accessVbaCfb(data)),
			toFileBytes: (): Buffer => {
				throw new MacroContainerError(
					'Access databases are read-only: Access renders and runs VBA from its compiled p-code tables, not from the source cache XLIDE reads.',
				);
			},
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
		vbaCfb: cached(() => Cfb.fromBytes(xlsx.readVbaProject())),
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
	throw new MacroContainerError(
		'Compound file without a recognizable Office host (no WordDocument, Workbook, or PowerPoint Document stream).',
	);
}

/** .doc / .xls: the file IS the CFB, so writing is re-serializing it. */
function wholeCfbContainer(outer: Cfb, kind: MacroContainerKind, description: string): MacroContainer {
	return {
		kind,
		writable: true,
		description,
		vbaCfb: (): Cfb => outer,
		toFileBytes: (cfb: Cfb): Buffer => cfb.toBytes(),
	};
}

function cached(build: () => Cfb): () => Cfb {
	let value: Cfb | undefined;
	return (): Cfb => {
		value ??= build();
		return value;
	};
}
