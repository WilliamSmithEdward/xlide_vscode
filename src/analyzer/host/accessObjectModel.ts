// The Access host object model: generated type-library metadata plus the
// host's injected globals. Access is READ-ONLY territory for XLIDE - the
// .accdb container is not OOXML and no engine write path exists - so this
// model serves the language surfaces: completion, hover, and honest
// diagnostics for VBA read out of an Access project.
//
// `CurrentDb` is deliberately absent from the globals: it returns a
// DAO.Database, a library this model does not carry, and mapping it to
// anything else would be a guess. Everything remains NON-exhaustive.

import type { HostObjectModel } from './excelObjectModel';
import { mergeHostConstants } from './excelObjectModel';
import { OFFICE_REFERENCE_ENUM_CONSTANTS } from './officeReferenceConstants';
import { accessReferenceData } from './accessObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getAccessObjectModel(): HostObjectModel {
	if (MODEL) {
		return MODEL;
	}
	// accessReferenceData() evaluates its metadata literals on first call,
	// so sessions that never touch an Access file never pay.
	const data = accessReferenceData();
	MODEL = {
		source: 'Microsoft Access 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn; shared Office reference enum constants',
		hostName: 'Access',
		types: data.types as HostObjectModel['types'],
		aliases: data.aliases as HostObjectModel['aliases'],
		// The shared Office library is auto-referenced in every Access VBA
		// project; Access's own names have no overlap with it (measured).
		constants: mergeHostConstants(OFFICE_REFERENCE_ENUM_CONSTANTS, data.constants),
		globals: {
			Application: 'Access.Application',
			DoCmd: 'Access.DoCmd',
			Forms: 'Access.Forms',
			Reports: 'Access.Reports',
			Modules: 'Access.Modules',
			Screen: 'Access.Screen',
			CurrentProject: 'Access.CurrentProject',
			CurrentData: 'Access.CurrentData',
			CodeContextObject: 'Access.Application',
		},
	};
	return MODEL;
}
