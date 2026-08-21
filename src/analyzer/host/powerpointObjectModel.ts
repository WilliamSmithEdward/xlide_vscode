// The PowerPoint host object model: generated type-library metadata plus the
// host's injected globals. PowerPoint has no document modules - its VBA lives
// in standard modules, classes and forms - so unlike Word there is no
// ThisDocument-style name to type; the globals are the application surface.
// Everything remains NON-exhaustive: offers and describes, never proves
// absence.

import type { HostObjectModel } from './excelObjectModel';
import { mergeHostConstants } from './excelObjectModel';
import { OFFICE_REFERENCE_ENUM_CONSTANTS } from './officeReferenceConstants';
import { MSFORMS_REFERENCE_ENUM_CONSTANTS } from './msformsReferenceMembers';
import { officeReferenceTypeData } from './officeReferenceTypes';
import { powerpointReferenceData } from './powerpointObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getPowerPointObjectModel(): HostObjectModel {
	if (MODEL) {
		return MODEL;
	}
	// powerpointReferenceData() evaluates its metadata literals on first
	// call, so sessions that never touch a PowerPoint file never pay.
	const data = powerpointReferenceData();
	MODEL = {
		source: 'Microsoft PowerPoint 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn; shared Office reference enum constants',
		hostName: 'PowerPoint',
		globalType: 'PowerPoint.Global',
		// The shared Office library's types, merged under the host's own so a
		// chain that lands on one (TextFrame2.TextRange -> Office.TextRange2)
		// keeps resolving. The host wins every shared name.
		types: { ...officeReferenceTypeData().types, ...data.types } as HostObjectModel['types'],
		aliases: { ...officeReferenceTypeData().aliases, ...data.aliases } as HostObjectModel['aliases'],
		// The shared Office library is auto-referenced in every PowerPoint VBA
		// project (msoTrue, msoShapeRectangle, ...); PowerPoint's own table
		// wins the shared chart-enum names (same values by measurement).
		constants: mergeHostConstants(OFFICE_REFERENCE_ENUM_CONSTANTS, MSFORMS_REFERENCE_ENUM_CONSTANTS, data.constants),
		globals: {
			Application: 'PowerPoint.Application',
			ActivePresentation: 'PowerPoint.Presentation',
			Presentations: 'PowerPoint.Presentations',
			ActiveWindow: 'PowerPoint.DocumentWindow',
			Windows: 'PowerPoint.DocumentWindows',
			SlideShowWindows: 'PowerPoint.SlideShowWindows',
		},
	};
	return MODEL;
}
