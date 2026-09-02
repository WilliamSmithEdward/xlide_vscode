// The VB6 host object model: the runtime's own objects and constants
// (VBRUN, read from the type library inside msvbvm60.dll) and the VB library
// (App, Screen, Printer, Clipboard, Global, Form and the intrinsic controls),
// transcribed from twinBASIC's package documentation because VB6's VB.OLB is
// not available to introspect. Every dump names its source, and every type
// carries it as provenance.
//
// This model offers and describes; it never proves a member absent
// (roadmap_vb6_support.md, Slice 3). A `VB` member the documentation marks
// "reserved for VB6 compatibility" is real VB6 that twinBASIC does not run,
// so it is carried with that note in its remarks for the oracle harness to
// read. VBA6 itself (msvbvm60.dll resource 1) is dumped as evidence but not
// modelled here: the analyzer's VBA runtime already answers those names for
// every host.

import type { HostObjectModel } from './excelObjectModel';
import { vb6ReferenceData } from './vb6ObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getVb6ObjectModel(): HostObjectModel {
	if (MODEL) {
		return MODEL;
	}
	// vb6ReferenceData() evaluates its metadata literals on first call, so
	// sessions that never open a .vbp never pay for them.
	const data = vb6ReferenceData();
	MODEL = {
		source: 'VBRUN from msvbvm60.dll resource 3 via scripts/dump-vb6-typelib.py; VB transcribed from twinBASIC documentation via scripts/transcribe-vb6-docs.mjs',
		hostName: 'VB6',
		// The Global object's members (App, Screen, Forms, Load, Unload,
		// LoadPicture ...) are reachable without qualification in every module.
		globalType: 'VB.Global',
		types: data.types as HostObjectModel['types'],
		aliases: data.aliases as HostObjectModel['aliases'],
		enums: data.enums,
		constants: data.constants,
		globals: {
			App: 'VB.App',
			Screen: 'VB.Screen',
			Printer: 'VB.Printer',
			Printers: 'VB.Printers',
			Clipboard: 'VB.Clipboard',
		},
	};
	return MODEL;
}
