// The PowerPoint host object model: generated type-library metadata plus the
// host's injected globals. PowerPoint has no document modules - its VBA lives
// in standard modules, classes and forms - so unlike Word there is no
// ThisDocument-style name to type; the globals are the application surface.
// Everything remains NON-exhaustive: offers and describes, never proves
// absence.

import type { HostObjectModel } from './excelObjectModel';
import {
	POWERPOINT_REFERENCE_ALIASES,
	POWERPOINT_REFERENCE_CONSTANTS,
	POWERPOINT_REFERENCE_TYPES,
} from './powerpointObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getPowerPointObjectModel(): HostObjectModel {
	MODEL ??= {
		source: 'Microsoft PowerPoint 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn',
		types: POWERPOINT_REFERENCE_TYPES as HostObjectModel['types'],
		aliases: POWERPOINT_REFERENCE_ALIASES as HostObjectModel['aliases'],
		constants: POWERPOINT_REFERENCE_CONSTANTS as HostObjectModel['constants'],
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
