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
import {
	ACCESS_REFERENCE_ALIASES,
	ACCESS_REFERENCE_CONSTANTS,
	ACCESS_REFERENCE_TYPES,
} from './accessObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getAccessObjectModel(): HostObjectModel {
	MODEL ??= {
		source: 'Microsoft Access 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn',
		types: ACCESS_REFERENCE_TYPES as HostObjectModel['types'],
		aliases: ACCESS_REFERENCE_ALIASES as HostObjectModel['aliases'],
		constants: ACCESS_REFERENCE_CONSTANTS as HostObjectModel['constants'],
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
