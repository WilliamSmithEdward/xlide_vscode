// Every host model XLIDE ships, for a caller that analyzes every host.
//
// Kept out of hostRegistry so that importing the analyzer does not import
// them: together they are more than half of its bundled size. The extension
// and its analysis worker call this once at load; an embedder that only
// analyzes Excel never imports this file.

import { getAccessObjectModel } from './accessObjectModel';
import { registerHostObjectModel } from './hostRegistry';
import { getPowerPointObjectModel } from './powerpointObjectModel';
import { getVb6ObjectModel } from './vb6ObjectModel';
import { getWordObjectModel } from './wordObjectModel';

/** Registers the Word, PowerPoint, Access and VB6 models. Safe to call again. */
export function registerBuiltInHostModels(): void {
	registerHostObjectModel('word', getWordObjectModel);
	registerHostObjectModel('powerpoint', getPowerPointObjectModel);
	registerHostObjectModel('access', getAccessObjectModel);
	registerHostObjectModel('vb6', getVb6ObjectModel);
}
