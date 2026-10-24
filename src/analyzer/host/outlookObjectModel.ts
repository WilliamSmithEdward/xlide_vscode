// The Outlook library a project can reference: generated type-library
// metadata (scripts/dump-scripting-typelibs.py, then
// generate-host-object-model.mjs outlook) for the VBA of an Excel, Word,
// PowerPoint or Access project that drives Outlook early bound. Selected by
// the project's reference to the library's GUID, never added implicitly, and
// not a host: XLIDE does not open Outlook's own VbaProject.OTM.
//
// Outlook's interfaces are dispatch-only and extensible: `o.NoSuchMember` on
// an Outlook.Application compiles in every host (measured in Excel, Word,
// PowerPoint and Access 16.0 through the VBE oracle). So these types offer
// and describe, and never prove a member absent.

import type { HostObjectModel } from './excelObjectModel';
import { outlookReferenceData } from './outlookObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getOutlookObjectModel(): HostObjectModel {
	return MODEL ??= {
		...outlookReferenceData(),
		source: 'Microsoft Outlook 16.0 Object Library 9.6 via pyVBAReference and the registered COM type library',
		hostName: 'Outlook',
		// The library marks Application APPOBJECT, and its members compile
		// bare in a project that references it: Session, CreateItem,
		// GetNamespace (measured in all four hosts).
		globalType: 'Outlook.Application',
		globals: {},
	};
}
