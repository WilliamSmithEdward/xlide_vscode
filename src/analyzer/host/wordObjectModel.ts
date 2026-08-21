// The Word host object model: generated type-library metadata plus the one
// thing a type library cannot say - which names the host injects into the
// bare global scope of every module.
//
// Word's injected globals are the members of the hidden `Global` interface
// (Word.Global in the library); the short list here is the everyday core,
// curated the way Excel's globals are, and each maps to a type the generated
// metadata actually carries (pinned by test). Everything remains
// NON-exhaustive: this model offers and describes, never proves absence.

import type { HostObjectModel } from './excelObjectModel';
import { mergeHostConstants } from './excelObjectModel';
import { OFFICE_REFERENCE_ENUM_CONSTANTS } from './officeReferenceConstants';
import { MSFORMS_REFERENCE_ENUM_CONSTANTS } from './msformsReferenceMembers';
import { officeReferenceTypeData } from './officeReferenceTypes';
import { wordReferenceData } from './wordObjectModelData';

let MODEL: HostObjectModel | undefined;

export function getWordObjectModel(): HostObjectModel {
	if (MODEL) {
		return MODEL;
	}
	// wordReferenceData() evaluates its metadata literals on first call, so
	// sessions that never touch a Word file never pay for them.
	const data = wordReferenceData();
	MODEL = {
		source: 'Microsoft Word 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn; shared Office reference enum constants',
		hostName: 'Word',
		globalType: 'Word.Global',
		// The shared Office library's types, merged under the host's own so a
		// chain that lands on one (TextFrame2.TextRange -> Office.TextRange2)
		// keeps resolving. The host wins every shared name.
		types: { ...officeReferenceTypeData().types, ...data.types } as HostObjectModel['types'],
		aliases: { ...officeReferenceTypeData().aliases, ...data.aliases } as HostObjectModel['aliases'],
		// Every Word VBA project auto-references the shared Office library, so
		// msoTrue and friends are legal everyday names; Word's own table wins
		// the shared chart-enum names (same values by measurement).
		constants: mergeHostConstants(OFFICE_REFERENCE_ENUM_CONSTANTS, MSFORMS_REFERENCE_ENUM_CONSTANTS, data.constants),
		globals: {
			Application: 'Word.Application',
			ActiveDocument: 'Word.Document',
			ThisDocument: 'Word.Document',
			Documents: 'Word.Documents',
			Selection: 'Word.Selection',
			ActiveWindow: 'Word.Window',
			Windows: 'Word.Windows',
			Options: 'Word.Options',
			Templates: 'Word.Templates',
			NormalTemplate: 'Word.Template',
			System: 'Word.System',
			Tasks: 'Word.Tasks',
			Dialogs: 'Word.Dialogs',
		},
	};
	return MODEL;
}
