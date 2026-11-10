/**
 * The language ids XLIDE gives documents, as `package.json` contributes them.
 * Code that means a language names it through these, so a bare `'vba'`
 * anywhere else is the VBA library (`VBA.Left`), or the diagnostic
 * collection's name, never a language.
 * `tests/languageIds.test.ts` holds both sides to that.
 */

/** A VBA file on disk: `.bas`, `.cls`, `.frm` and the rest. */
export const VBA_LANGUAGE_ID = 'vba';

/** A module XLIDE serves from inside an Office file or a VB6 project. */
export const XLIDE_VBA_LANGUAGE_ID = 'xlide-vba';

/** A UserForm's markup (`.form`). */
export const XML_LANGUAGE_ID = 'xml';

/** The line that opens a Markdown code block coloured as VBA. */
export const VBA_CODE_FENCE = '```' + VBA_LANGUAGE_ID;
