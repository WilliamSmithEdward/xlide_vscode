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
		source: 'Microsoft Word 16.0 Object Library via pyVBAReference; enriched from Microsoft Learn',
		hostName: 'Word',
		types: data.types as HostObjectModel['types'],
		aliases: data.aliases as HostObjectModel['aliases'],
		constants: data.constants as HostObjectModel['constants'],
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
