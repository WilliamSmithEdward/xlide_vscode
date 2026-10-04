# Module header documentation work

`extractModuleHeaderDoc` built a physical-line table for the complete source before deciding whether its first lines held header documentation. Every edited symbol snapshot could allocate tens of thousands of unnecessary line records/substrings in the 969,296-character ROneCOne class. Stream lines and stop at the first header boundary or code line instead. Existing start-offset behavior, CR/LF/CRLF handling and module-versus-member/directive documentation ownership are preserved.

Work-count tests verify four physical-line reads for a documented header above a 26,000-line body, and bounded header reads while building symbols before/after a body edit. Offset, XML content and boundary cases are covered alongside existing documentation and editor regression suites.

An alternating 21-batch comparison (ten calls per batch) on the private class source measured median header lookup of 1.365 ms before and 0.00087 ms after; slowest batch averages were 2.024 and 0.01937 ms. This is the no-header lookup that previously scanned the entire class. It measures a component; it does not establish the same factor for total editor latency.

Three unprofiled VS Code 1.139.1 workload repeats passed all five actual-workbook cases: 144 valid typing/Backspace pairs and 24 missed-prefix member recovery/hover cycles. Typing medians were 2.44/2.14/2.35 ms, Backspace medians were 1.88/1.79/2.20 ms, and Backspace maxima were 98.87/25.96/24.18 ms. Heartbeat maxima were 64.05/30.00/45.33 ms. Fresh declaration hover took 34.06/37.39/29.41 ms. Member completion after deletion reached 37.16 ms. These separate runs include editor/background activity and do not measure painting or prove the disappearance of all future outliers. The preceding combined-code runs recorded a 304.80 ms Backspace maximum.

Compilation and 179 focused documentation/editor tests passed. The full suite passed 14,398 tests across 729 files, including private-workbook opt-in probes; 24 tests were intentionally skipped. No original source, workbook or CPU profile is committed. See INTERMITTENT_TYPING_STALLS.md for reproduction, command timing limits and the remaining stall investigation.
