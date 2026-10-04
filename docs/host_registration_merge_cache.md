# Refresh merged models after host provider registration

`registerHostObjectModel` replaced the provider in its token registry but left `MERGED_BY_KEY` unchanged. Once two registered host models had been merged, replacing either provider made single-token lookups see the new model while merged lookups still returned old types, aliases, globals, constants, enums, signatures and provenance.

Registration now invalidates cached combinations containing that exact token. The next lookup merges the current providers, retaining first-token precedence and secondary-enum library labels. Unaffected combinations retain their cached object identity. Re-registering the same factory also refreshes affected combinations, so a factory that now supplies a new model can signal that change. Previously returned model objects remain unchanged.

Six new tests cover complete metadata after primary/secondary replacement, all cached orderings and repeated tokens, same-factory refresh, unaffected-cache identity, newly registered tokens and absent/Excel defaults. Four cases fail on the baseline. All 80 focused registry/interoperability/host-seam tests and type checking passed. This is a cache correctness repair; no timing or heap-byte improvement is claimed. Registration scans existing cache keys, while normal lookups retain their cache behavior.

Full repository suite: 753 files passed, seven skipped; 14,695 tests passed, 33 skipped (77.19 seconds).
