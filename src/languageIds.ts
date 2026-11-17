/**
 * The ids of the languages XLIDE contributes in `package.json`. Code that
 * means a language names it through these, which leaves a bare `'vba'` in the
 * source to mean the VBA library (`VBA.Left`). It imports nothing, so the
 * analyzer can use it as well as the editor layer.
 * `tests/languageIds.test.ts` holds the values to `package.json`.
 */

/** A VBA file on disk: `.bas`, `.cls`, `.frm` and the other module files. */
export const VBA_LANGUAGE_ID = 'vba';

/** A module XLIDE serves from inside an Office file or a VB6 project. */
export const XLIDE_VBA_LANGUAGE_ID = 'xlide-vba';

/** A form's markup (`.form`), which XLIDE shows as XML. */
export const FORM_MARKUP_LANGUAGE_ID = 'xml';

/** The line that opens a Markdown code block highlighted as VBA. */
export const VBA_CODE_FENCE = `\`\`\`${VBA_LANGUAGE_ID}`;
