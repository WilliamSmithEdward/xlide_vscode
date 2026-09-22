// The one signal for "this file is a container XLIDE reads, and there is no
// VBA project inside it".
//
// Every container can be in that state, and reaching it is ordinary: Excel
// writes no `xl/vbaProject.bin` part at all into a workbook saved as .xlsm
// before the first macro exists, and a .doc, a .ppt or an .accdb with no code
// is the same shape. It is a STATE, not a failure - a file XLIDE read
// perfectly well that happens to hold nothing yet - so the surfaces that list
// code answer "nothing yet" instead of the "Load failed - click to retry" a
// file XLIDE genuinely could not read gets.
//
// It lives in its own module because all four container readers raise it and
// macroContainer, the seam callers catch it from, imports all four: putting
// the class beside MacroContainerError would make that import circular.

/** Raised where a readable container simply has no VBA project in it yet. */
export class NoVbaProjectError extends Error {
	/**
	 * What the file is, in macroContainer's own words ("an Excel workbook"),
	 * so a caller can name the container in a sentence of its own.
	 */
	readonly containerDescription: string;

	constructor(containerDescription: string) {
		super(`No VBA project: this file is ${containerDescription} with no VBA in it yet.`);
		this.name = 'NoVbaProjectError';
		this.containerDescription = containerDescription;
	}
}
