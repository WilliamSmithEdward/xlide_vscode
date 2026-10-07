# Host semantic member lookup performance

Host member painting previously linearly searched each resolved member surface. Its collector now owns a fresh instance of the existing member-surface cache per pass, and cached callers reuse the existing case-insensitive first-match member index. Uncached API calls retain their original linear path, including cheap first-member lookups. Cache lifetime stays within the collection pass; no global source cache or dependency is added.

## Reproduce

```powershell
node scripts/benchmark-semantic-host-members.mjs --baseline=739a23a3 > before.json
node scripts/benchmark-semantic-host-members.mjs > after.json
```

Node v24.18.0 on Windows, Ryzen 7 9800X3D. Four isolated processes ran sequentially in baseline/fixed/fixed/baseline order. Each configuration has 15 timed rounds after three warm-ups; no full suite or corpus job overlapped timing.

## Method and scope

- Runs the actual collectHostMemberMethodTokens consumer, including receiver resolution, parsing/cache access and exact token creation. Synthetic ordinary host metadata has 5 or 1,000 method/property members; a second type provides a globally known name missing from the receiver type. Metadata construction and complete-output assertions/hashes are outside the timer.
- First/last/missing/unknown queries use a declared local As Range. Combined document queries use a Sheet1 code name and document project type mapped to the synthetic host Range. This exercises the real combined surface merge; it does not execute Office or measure editor rendering.
- Cached-source passes reuse the source and host model, while fresh-source passes change a trailing comment to bypass parser/source caches. Both get fresh collector-owned member caches. Model-cold-first uses a fresh host model every round, with host-model indexing inside the timer.
- Every timed result matches its fixture’s complete token array. All 42 configurations have matching complete result hashes across the four processes. A separate immutable metadata getter regression shows 1,000 last-member references reading 1,000 names once, versus 1,000,000 on baseline.

## Results

Times are milliseconds. Ranges span the two process medians or p95s for each implementation.

| Members | Query | References | Source/model scope | Baseline median | Fixed median | Baseline p95 | Fixed p95 |
| ---: | --- | ---: | --- | ---: | ---: | ---: | ---: |
| 5 | first | 1 | cached-source | 0.010–0.011 | 0.010–0.010 | 0.029–0.030 | 0.028–0.032 |
| 5 | first | 1 | fresh-source | 0.037–0.038 | 0.035–0.041 | 0.075–0.076 | 0.076–0.077 |
| 5 | first | 1000 | cached-source | 0.733–0.812 | 0.578–0.622 | 1.277–3.513 | 1.057–1.159 |
| 5 | first | 1000 | fresh-source | 1.803–1.973 | 1.836–1.949 | 2.892–3.450 | 3.042–3.101 |
| 5 | last | 1 | cached-source | 0.001–0.001 | 0.002–0.002 | 0.001–0.003 | 0.002–0.003 |
| 5 | last | 1 | fresh-source | 0.006–0.007 | 0.007–0.007 | 0.009–0.011 | 0.008–0.008 |
| 5 | last | 1000 | cached-source | 0.454–0.454 | 0.360–0.376 | 0.634–0.714 | 0.411–0.581 |
| 5 | last | 1000 | fresh-source | 1.221–1.231 | 1.139–1.297 | 1.452–1.478 | 1.408–2.093 |
| 5 | missing | 1 | cached-source | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.002–0.003 |
| 5 | missing | 1 | fresh-source | 0.006–0.007 | 0.006–0.010 | 0.009–0.018 | 0.021–0.042 |
| 5 | missing | 1000 | cached-source | 0.401–0.408 | 0.327–0.347 | 0.575–0.747 | 0.479–0.486 |
| 5 | missing | 1000 | fresh-source | 1.178–1.293 | 1.265–1.523 | 1.867–1.871 | 1.879–2.031 |
| 5 | combined-last | 1 | cached-source | 0.005–0.005 | 0.005–0.005 | 0.017–0.018 | 0.016–0.022 |
| 5 | combined-last | 1 | fresh-source | 0.008–0.010 | 0.011–0.012 | 0.016–0.021 | 0.017–0.020 |
| 5 | combined-last | 1000 | cached-source | 2.126–2.160 | 1.836–1.919 | 2.523–2.661 | 2.218–2.281 |
| 5 | combined-last | 1000 | fresh-source | 2.984–3.066 | 2.660–2.670 | 3.299–3.366 | 3.130–3.301 |
| 5 | unknown | 1 | cached-source | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 |
| 5 | unknown | 1 | fresh-source | 0.005–0.006 | 0.005–0.006 | 0.018–0.020 | 0.017–0.024 |
| 5 | unknown | 1000 | cached-source | 0.032–0.032 | 0.032–0.032 | 0.032–0.033 | 0.032–0.037 |
| 5 | unknown | 1000 | fresh-source | 0.761–0.796 | 0.803–0.826 | 1.100–1.109 | 1.109–1.232 |
| 5 | model-cold-first | 1 | cached-source | 0.002–0.002 | 0.003–0.003 | 0.003–0.003 | 0.004–0.004 |
| 1000 | first | 1 | cached-source | 0.001–0.001 | 0.045–0.046 | 0.001–0.002 | 0.047–0.058 |
| 1000 | first | 1 | fresh-source | 0.005–0.005 | 0.050–0.050 | 0.010–0.011 | 0.066–0.067 |
| 1000 | first | 1000 | cached-source | 0.463–0.498 | 0.435–0.463 | 0.508–0.835 | 0.453–0.507 |
| 1000 | first | 1000 | fresh-source | 1.184–1.223 | 1.123–1.174 | 1.487–1.568 | 1.439–1.485 |
| 1000 | last | 1 | cached-source | 0.008–0.008 | 0.045–0.056 | 0.009–0.009 | 0.053–0.070 |
| 1000 | last | 1 | fresh-source | 0.013–0.013 | 0.049–0.050 | 0.016–0.016 | 0.051–0.052 |
| 1000 | last | 1000 | cached-source | 7.949–8.075 | 0.391–0.397 | 8.274–8.587 | 0.420–1.731 |
| 1000 | last | 1000 | fresh-source | 8.816–8.874 | 1.059–1.103 | 10.405–11.530 | 1.441–1.613 |
| 1000 | missing | 1 | cached-source | 0.007–0.007 | 0.045–0.045 | 0.008–0.009 | 0.046–0.046 |
| 1000 | missing | 1 | fresh-source | 0.012–0.012 | 0.049–0.050 | 0.012–0.013 | 0.065–1.268 |
| 1000 | missing | 1000 | cached-source | 6.941–6.955 | 0.367–0.393 | 7.117–8.160 | 0.400–0.666 |
| 1000 | missing | 1000 | fresh-source | 7.800–8.044 | 1.087–1.117 | 7.961–8.805 | 1.569–2.106 |
| 1000 | combined-last | 1 | cached-source | 0.053–0.053 | 0.102–0.129 | 0.066–0.227 | 0.121–0.161 |
| 1000 | combined-last | 1 | fresh-source | 0.058–0.058 | 0.116–0.134 | 0.069–0.071 | 0.128–0.153 |
| 1000 | combined-last | 1000 | cached-source | 53.830–54.936 | 1.833–2.123 | 57.169–57.786 | 1.964–3.288 |
| 1000 | combined-last | 1000 | fresh-source | 54.591–55.794 | 2.703–2.727 | 56.716–56.815 | 3.316–3.737 |
| 1000 | unknown | 1 | cached-source | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 |
| 1000 | unknown | 1 | fresh-source | 0.005–0.007 | 0.005–0.005 | 0.005–0.011 | 0.007–0.013 |
| 1000 | unknown | 1000 | cached-source | 0.032–0.032 | 0.032–0.034 | 0.032–0.032 | 0.036–0.038 |
| 1000 | unknown | 1000 | fresh-source | 0.787–0.830 | 0.794–0.806 | 1.317–1.327 | 1.384–1.583 |
| 1000 | model-cold-first | 1 | cached-source | 0.092–0.145 | 0.137–0.138 | 0.157–0.162 | 0.156–0.160 |

At 1,000 members/references, cached last-member collection improves from 7.949–8.075 ms to 0.391–0.397 ms; combined-document collection improves from 53.830–54.936 ms to 1.833–2.123 ms. Fresh-source gains remain substantial for these workloads. Small/early-hit/fresh/unknown controls are mixed or similar.

Index construction costs remain visible: a single first-member query on a warm 1,000-member model is 0.001 ms versus 0.045–0.046 ms, and a single combined-document query is 0.053 ms versus 0.102–0.129 ms. The model-cold control overlaps across processes. These bounded cold costs are included above; no universal whole-provider or UI speedup is claimed.

## Validation

Seven new regressions include two baseline work-count failures and five controls for no-reference passes, cheap uncached calls, changed code names, first case-insensitive matches and project shadowing across passes. 138 focused tests and type checks pass. On main baseline 739a23a3 plus this fix, the full suite passes 649 files with 13,193 tests passed and 13 skipped.

Complete semantic token arrays match for 132,096 collections: all 8,256 oracle sources across Excel, Word, PowerPoint and Access with four contexts (default, host-specific document/code-name, project type shadowing and implicit control shadowing). Full-source cached tokens and significant cached statement tokens, including their arrays, are frozen in both bundles; the run completed without uncaught exceptions. This does not claim a frozen AST or exhaustive repository coverage.
