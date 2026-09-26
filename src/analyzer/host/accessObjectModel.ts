// The Access host object model: generated type-library metadata plus the
// host's injected globals. It serves the language surfaces: completion,
// hover, and honest diagnostics for the VBA of an Access project.
//
// `CurrentDb` is deliberately absent from the globals: it returns a
// DAO.Database, a library whose TYPES this model does not carry, and mapping
// it to anything else would be a guess. The DAO library's CONSTANTS are here
// (issue #103): every new database references it, and dbFailOnError,
// dbOpenDynaset and dbText are everyday Access names. Access-library types are
// exhaustive once the hidden members are merged (issue #127) and prove a
// member absent where the interface is closed; the controls are, Form is not.

import type { HostObjectModel } from './excelObjectModel';
import { hostBoundOfficeTypes, mergeHostConstants } from './excelObjectModel';
import { officeReferenceTypeData } from './officeReferenceTypes';
import { OFFICE_REFERENCE_ENUM_CONSTANTS, OFFICE_REFERENCE_ENUMS } from './officeReferenceConstants';
import { MSFORMS_REFERENCE_ENUM_CONSTANTS } from './msformsReferenceMembers';
import { DAO_REFERENCE_ENUM_CONSTANTS, DAO_REFERENCE_ENUMS } from './daoReferenceConstants';
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
		source: 'Microsoft Access 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn; shared Office reference enum constants; DAO (Access Database Engine 16.0) enum constants',
		hostName: 'Access',
		// The shared Office library's types, merged under the host's own so a
		// chain that lands on one (TextFrame2.TextRange -> Office.TextRange2)
		// keeps resolving. The host wins every shared name.
		types: { ...hostBoundOfficeTypes('Access.Application'), ...data.types } as HostObjectModel['types'],
		aliases: { ...officeReferenceTypeData().aliases, ...data.aliases } as HostObjectModel['aliases'],
		// The shared Office library is auto-referenced in every Access VBA
		// project; Access's own names have no overlap with it (measured).
		// The host's own enumerations win a shared name, as its constants do.
		// DAO is auto-referenced too (issue #103); the host's own names win
		// any it shares, and Access's own table shares none (measured).
		enums: { ...OFFICE_REFERENCE_ENUMS, ...DAO_REFERENCE_ENUMS, ...data.enums },
		constants: mergeHostConstants(
			OFFICE_REFERENCE_ENUM_CONSTANTS,
			MSFORMS_REFERENCE_ENUM_CONSTANTS,
			DAO_REFERENCE_ENUM_CONSTANTS,
			data.constants,
		),
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
