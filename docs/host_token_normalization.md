# Normalize host tokens before merged-model selection

Single-host `hostObjectModelForToken` trims and lowercases host tokens, but multi-host `hostObjectModelForTokens` filtered its input against case-sensitive registry keys first. Adding referenced hosts could therefore silently omit the project's own host or referenced libraries when tokens contained casing differences or whitespace. The selected globals, ambiguous type precedence and available library types could change.

Normalize each token before registry filtering and cache-key construction, preserving token order and duplicate behavior. Equivalent normalized inputs share the existing merged model cache. Absent/Excel defaults and unknown-token behavior remain unchanged; caller arrays are not modified.

Eleven new tests cover uppercase, whitespace and mixed tokens across two host orderings and a three-host merge; exact cache identity, host/type precedence, frozen inputs, analyzer option integration, defaults, duplicates and unknown hosts. Ten cases fail on the baseline. Type checking and 97 focused host tests passed. This repairs consistent model selection; no timing or heap-byte performance claim.

Full repository suite: 754 files passed, seven skipped; 14,710 tests passed, 33 skipped (89.43 seconds).
