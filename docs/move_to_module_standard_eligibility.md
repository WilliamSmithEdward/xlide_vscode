# Move to Module standard-module eligibility

The refactor's standard-module contract was not enforced. Its command offered every module and failed to pass host roles it already read. The analyzer could remove a procedure from a class, document or form or insert it into one, changing object context and visibility behavior.

MoveToModuleInput now accepts optional moduleKinds metadata. Names are compared case-insensitively; missing entries keep the legacy standard-module default. Source and destination must both be standard modules; refusal produces no edits. The actual command passes existing host metadata and filters destinations to standard modules. The shared reference-binding index retains each caller's actual module role. This requires no additional bridge reads or source parses for the eligibility check.

Validation against 8381e9ba:

- 18 analyzer eligibility failures and six independently expected full-output controls across LF/CRLF/CR before the fix.
- Seven actual-command baseline failures and one unchanged legacy control: unsafe object moves reach applyEdit/writeProjectModule, and the picker includes invalid destinations.
- All 120 focused checks across eight files pass with the fix; after rebasing onto main f006e788, types and all 123 affected checks across nine files pass. Command tests cover class, document, userform, usercontrol, propertypage, designer, accessform and accessreport module types; refused moves perform no workspace edit, project write or refresh.
- Full unit suite passes: 718 files, six skipped; 14,348 tests passed, 30 skipped. No failures.
- Type checking passes. Frozen-input corpus comparison against parent 8381e9ba after integration preserves 16,500 full results and outputs from 8,256 sources and 72 generated recursive moves. Corpus calls omit role metadata and are compatibility evidence; explicit metadata behavior is covered by the dedicated engine/command tests.

This fix enforces endpoint eligibility, not private-procedure visibility or destination name capture. Those remain separate audits. No runtime speedup is claimed.
