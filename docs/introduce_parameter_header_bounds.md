# Introduce Parameter logical-header insertion

Introduce Parameter sliced the entire procedure twice and searched for the first closing parenthesis. Without a parameter list, a closing parenthesis in a body expression or header comment became the insertion point. With a required array parameter, the first closing parenthesis ended the array suffix rather than the outer list. Both cases wrote invalid code.

One helper now uses the parsed procedure name/type-suffix spans, the existing logical-header boundary helper, and the cached token stream. An absent list receives parentheses after the name/suffix. An existing list receives the new parameter at its balanced outer close, bounded by the header so a body token cannot become the insertion point. String/comment token contents do not count as delimiters. The helper supplies the edit and its punctuation together, removing the two duplicated whole-procedure searches and the obsolete Span import. An unreliable header is refused before edits are produced.

Validation against c3b49c07b1e13aad5118d250947f22d55aa5bef7:

- Twelve regression outputs fail before the fix; three simple-list controls pass. All fifteen pass after the fix across LF/CRLF/CR, a body expression containing a quoted closing parenthesis, a header comment containing one, a required array parameter, and a continued array-parameter header. Expected complete module output includes updated callers, unchanged body/comment text and removed local declaration/initializer.
- Types and 147 focused tests across seven files pass, including command binding, cross-module calls, recursive receivers, property eligibility and binding work counts.

This is a correctness and duplication repair, not a measured end-to-end performance claim. Other signature evolution concerns such as Optional/ParamArray ordering and AddressOf callers remain separate audit candidates; this change does not establish completion of the refactor surface.

Full suite: 732 files passed, seven skipped; 14,497 tests passed, 32 skipped. No failures.
