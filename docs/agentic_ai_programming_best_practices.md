# Programming Best Practices for Agentic AI Coding

**Version:** 2026-06-10  
**Scope:** language-agnostic programming practices for human engineers and agentic AI coding assistants.  
**Purpose:** give a coding agent clear, enforceable rules for producing maintainable, reviewable, testable, secure software.

---

## Qualification rule

A practice is included in the **Research-Gated Practices** section only when the same core concept is supported by at least **three independent sources** in the source matrix.

The **User-Mandated Design Principles** section is included by request and is **not research-gated**.

This guide is intentionally language-agnostic. Apply the rules through the conventions, frameworks, and tooling already present in the repository.

---

## How an agent should use this file

Treat this as an operating policy for software changes.

Priority order:

1. The user’s explicit task.
2. Repository-specific instructions such as `AGENTS.md`, `CLAUDE.md`, `.github/copilot-instructions.md`, `CONTRIBUTING.md`, or local docs.
3. Existing project architecture and style.
4. This guide.
5. General language or framework conventions.

When instructions conflict, preserve correctness, security, and maintainability. Prefer the smallest coherent change that satisfies the task.

---

## Agentic coding loop

For every task, follow this loop.

### 1. Understand

- Restate the goal internally in concrete terms.
- Inspect the relevant files before editing.
- Identify the existing architecture, naming style, test strategy, and build commands.
- Find the narrowest change boundary.
- Note uncertainty instead of inventing facts.

### 2. Plan

- Pick the smallest coherent implementation path.
- Identify the validation commands before coding.
- Avoid speculative rewrites.
- Avoid adding dependencies unless the benefit is clear and durable.
- Prefer project-native patterns over personal preferences.

### 3. Change

- Keep the diff focused.
- Preserve existing behavior unless the task explicitly changes it.
- Add or update tests with the code.
- Update docs when behavior, setup, APIs, or architecture change.
- Do not hardcode secrets, local paths, credentials, or machine-specific values.

### 4. Validate

Run the strongest available checks, in this order when applicable:

1. Targeted unit tests for changed behavior.
2. Type checks.
3. Lint or formatting checks.
4. Build or compile check.
5. Integration or smoke tests.
6. Security or dependency checks.

If a check cannot be run, say why and provide the best next validation step.

### 5. Report

In the final response or commit/PR description, include:

- What changed.
- Why it changed.
- Files touched.
- Validation commands and results.
- Known risks or follow-up work.

---

# Research-Gated Practices

## RG-01 — Make small, focused changes

**Core rule:** One task should produce one coherent change. Split unrelated work.

**Agent behavior:**

- Keep each patch focused on a single goal.
- Avoid mixing formatting, refactoring, dependency updates, and feature work unless they are required for the same change.
- Prefer a series of small safe edits over one large risky edit.
- Include enough context in the summary for a reviewer or future agent to understand the change.

**Acceptance criteria:**

- The diff has one dominant purpose.
- Each touched file is explainable.
- Tests or validation match the changed surface.
- The change can be reverted without removing unrelated work.

**Bad smells:**

- “While I was here” edits.
- Broad rewrites for a narrow bug.
- Large formatting churn mixed with logic changes.
- Multiple unrelated bug fixes in one patch.

---

## RG-02 — Require reviewable code before merge

**Core rule:** Code should be easy for another person or agent to review before it is merged.

**Agent behavior:**

- Self-review the diff before presenting it.
- Check design, functionality, complexity, names, tests, and security implications.
- Make assumptions explicit.
- Prefer clear code over clever code.
- Do not hide risky changes inside mechanical edits.

**Acceptance criteria:**

- A reviewer can understand the purpose quickly.
- The implementation is simpler than plausible alternatives.
- Test evidence is available.
- Risky tradeoffs are called out.

**Bad smells:**

- No explanation of why the change exists.
- Cleverness without a measurable benefit.
- Hidden behavior changes.
- Test gaps around modified logic.

---

## RG-03 — Validate every meaningful change

**Core rule:** Every non-trivial change needs deterministic validation.

**Agent behavior:**

- Identify validation commands before editing.
- Run targeted tests first.
- Run build, type, lint, and smoke checks when available.
- Capture exact commands and outcomes.
- If validation cannot run, explain the blocker and name the next best check.

**Acceptance criteria:**

- Validation directly exercises the changed behavior.
- Failures are investigated, not ignored.
- Known skipped checks are disclosed.
- The final report includes command evidence.

**Bad smells:**

- “Looks good” without running checks.
- Only testing unrelated paths.
- Ignoring flaky or failing tests without explanation.
- Relying on manual inspection for behavior that can be tested.

---

## RG-04 — Treat tests as behavior contracts

**Core rule:** Tests should specify important behavior and protect against regressions.

**Agent behavior:**

- Add tests for new behavior and bug fixes.
- Prefer fast, isolated, repeatable, self-checking tests.
- Cover edge cases, failure paths, and public contracts.
- Keep unit tests small and broad-stack tests purposeful.
- Do not over-mock behavior that should be verified end-to-end.

**Acceptance criteria:**

- Tests fail before the fix when practical.
- Tests pass after the fix.
- Tests are deterministic.
- Tests document the intended behavior.

**Bad smells:**

- Tests that require hidden local state.
- Tests that pass without asserting anything meaningful.
- Huge integration tests for simple pure logic.
- Mocking the system so heavily that no real behavior is checked.

---

## RG-05 — Prefer simplicity over unnecessary generality

**Core rule:** Solve the actual problem cleanly. Do not build speculative systems for imagined future needs.

**Agent behavior:**

- Choose the simplest design that satisfies current requirements.
- Remove dead paths and unused abstractions.
- Avoid speculative extension points.
- Use abstraction only when it reduces real duplication, isolates real volatility, or clarifies the model.
- Challenge complexity that has no clear payoff.

**Acceptance criteria:**

- The solution is understandable without a long explanation.
- Each abstraction has a current reason to exist.
- The code avoids needless branching and configuration.
- The implementation does not predict future requirements without evidence.

**Bad smells:**

- Factories, strategies, adapters, or plugins with one implementation and no current need.
- Parameters added only because “maybe later.”
- Code paths that cannot be reached.
- Complex generic code replacing simple direct logic.

---

## RG-06 — Use clear names and consistent style

**Core rule:** Code should communicate intent through names, structure, and consistent conventions.

**Agent behavior:**

- Follow the existing repository style.
- Use formatters and linters when present.
- Name variables, functions, types, files, and modules after their role in the domain.
- Prefer explicit intent over abbreviations.
- Keep naming consistent across related concepts.

**Acceptance criteria:**

- Names explain purpose, not implementation trivia.
- Similar things are named similarly.
- Formatting is automated or consistent with nearby code.
- A reader can skim the code and infer the flow.

**Bad smells:**

- Generic names such as `data`, `manager`, `helper`, `processor`, or `thing` without domain meaning.
- Multiple names for the same concept.
- New style conventions introduced in one file only.
- Comments needed because names are vague.

---

## RG-07 — Define explicit contracts and stable interfaces

**Core rule:** Boundaries between modules, services, APIs, and tools should be explicit and testable.

**Agent behavior:**

- Document public inputs, outputs, errors, and invariants.
- Use schemas, interface definitions, or contract tests where appropriate.
- Preserve existing contracts unless the task intentionally changes them.
- Version or migrate breaking changes deliberately.
- Avoid leaking internal implementation details through public interfaces.

**Acceptance criteria:**

- Consumers can understand how to call the interface.
- Invalid input behavior is defined.
- Error behavior is predictable.
- Contract changes are tested and documented.

**Bad smells:**

- Public APIs inferred only from implementation.
- Silent breaking changes.
- Multiple inconsistent shapes for the same concept.
- Undocumented magic fields or status values.

---

## RG-08 — Treat documentation as part of the deliverable

**Core rule:** If behavior, setup, architecture, commands, or APIs change, documentation should change with the code.

**Agent behavior:**

- Update README, usage docs, comments, runbooks, examples, or architecture notes when relevant.
- Keep docs close to the code they describe when possible.
- Prefer executable examples and exact commands.
- Remove stale docs during related changes.
- Record significant architectural decisions.

**Acceptance criteria:**

- A new contributor or agent can build, test, and use the changed feature.
- Public behavior is documented.
- Important decisions have context and consequences.
- Docs do not contradict the implementation.

**Bad smells:**

- Changed commands with unchanged README.
- New configuration with no explanation.
- Architecture changes hidden only in code.
- Comments that describe old behavior.

---

## RG-09 — Practice version-control hygiene

**Core rule:** The repository history should help future debugging, review, and rollback.

**Agent behavior:**

- Group related changes together.
- Separate unrelated changes.
- Write clear summaries explaining what and why.
- Avoid committing generated noise unless required.
- Keep rollback boundaries clean.

**Acceptance criteria:**

- Each commit or patch has a clear purpose.
- History can answer why a change happened.
- Reverting one change does not remove unrelated fixes.
- Generated files are handled according to project policy.

**Bad smells:**

- “misc fixes” commits.
- Formatting changes mixed with behavior changes.
- Large unrelated snapshots.
- Commit messages that only repeat file names.

---

## RG-10 — Manage dependencies and supply-chain risk

**Core rule:** Dependencies are part of the system’s attack surface and maintenance burden.

**Agent behavior:**

- Add dependencies only when they clearly beat local implementation.
- Prefer maintained, widely reviewed, appropriately licensed packages.
- Use lockfiles or equivalent reproducibility controls where applicable.
- Keep dependencies updated through automated or regular review.
- Scan for known vulnerable dependencies when tooling exists.
- Remove unused dependencies.

**Acceptance criteria:**

- Each dependency has a clear purpose.
- Known vulnerabilities are addressed or explicitly accepted with rationale.
- Dependency changes are isolated and reviewable.
- Builds are reproducible enough for the project’s risk level.

**Bad smells:**

- New package for trivial code.
- Unpinned or floating critical dependencies where reproducibility matters.
- Ignored vulnerability alerts.
- Abandoned packages in sensitive paths.

---

## RG-11 — Keep configuration and secrets out of source code

**Core rule:** Config should be externalized; secrets should never be hardcoded or committed.

**Agent behavior:**

- Use environment variables, config files, secret managers, or platform-native configuration.
- Never write API keys, tokens, passwords, private keys, or credentials into source.
- Add examples such as `.env.example` without real secrets.
- Redact secrets from logs, errors, tests, screenshots, and generated artifacts.
- Prefer automated secret scanning when available.

**Acceptance criteria:**

- Secrets are not present in code, tests, docs, logs, or fixtures.
- Local and production config can differ without code changes.
- Required configuration is documented.
- Accidental secret exposure is detectable.

**Bad smells:**

- `password = "..."` or `apiKey = "..."` in code.
- Machine-specific absolute paths.
- Production values in test fixtures.
- Logging full request headers or environment dumps.

---

## RG-12 — Use secure-by-design defaults

**Core rule:** Security should be built into normal design, not patched on after the fact.

**Agent behavior:**

- Validate untrusted input at trust boundaries.
- Encode or sanitize output for the target context.
- Use least privilege for files, network, credentials, database access, and service permissions.
- Prefer centralized authentication, authorization, validation, logging, and error-handling patterns.
- Avoid exposing stack traces or internal details to users.
- Make the safe path the default path.

**Acceptance criteria:**

- Untrusted input has a defined validation path.
- Sensitive operations require explicit authorization.
- Error responses are useful but do not leak internals.
- Security-relevant logic is tested or otherwise verified.

**Bad smells:**

- Ad hoc validation scattered across handlers.
- Catch-all authorization bypasses.
- Internal exception details returned to users.
- Dangerous defaults that rely on callers to opt into safety.

---

## RG-13 — Build observability into important flows

**Core rule:** Production behavior should be diagnosable from emitted signals.

**Agent behavior:**

- Add structured logs, metrics, traces, or events around important boundaries.
- Include correlation IDs or request context where appropriate.
- Log decisions, failures, retries, and external interactions at useful levels.
- Do not log secrets or excessive personal data.
- Prefer actionable signals over noisy logs.

**Acceptance criteria:**

- A maintainer can diagnose common failures without reproducing locally.
- Error paths emit enough context to investigate.
- Logs are structured enough for search and aggregation where applicable.
- Sensitive data is redacted.

**Bad smells:**

- Silent failure paths.
- Logs that say only “failed.”
- High-volume debug noise in normal operation.
- Secrets, tokens, or raw credentials in logs.

---

## RG-14 — Refactor safely and deliberately

**Core rule:** Refactoring should improve internal structure without changing external behavior.

**Agent behavior:**

- Separate refactoring from feature work when possible.
- Preserve public behavior unless the task says otherwise.
- Refactor in small steps.
- Add tests before risky refactors when coverage is weak.
- Explain why the refactor reduces complexity, duplication, coupling, or risk.

**Acceptance criteria:**

- Existing behavior remains intact.
- Tests or other validation protect the changed area.
- The refactor has a clear maintenance benefit.
- The diff is reviewable.

**Bad smells:**

- “Cleanup” that changes behavior accidentally.
- Large rewrites without tests.
- New abstractions that make the code harder to follow.
- Mixing unrelated refactors with feature changes.

---

## RG-15 — Use CI as a quality gate

**Core rule:** Automated checks should run consistently before code is considered mergeable.

**Agent behavior:**

- Preserve and use existing CI workflows.
- Add checks when introducing new validation needs.
- Keep feedback fast enough that contributors use it.
- Fail builds on meaningful test, lint, type, build, or security failures.
- Avoid flaky checks; quarantine or fix them.

**Acceptance criteria:**

- The same checks run predictably for contributors.
- CI catches common regression classes.
- Failed checks block merge or require explicit exception.
- Pipeline output is actionable.

**Bad smells:**

- Manual-only release validation.
- CI that is routinely red but ignored.
- Slow pipelines with no targeted fast path.
- Tests that pass locally but cannot run in CI.

---

## RG-16 — Record important architecture decisions

**Core rule:** Significant design decisions should leave a durable decision record.

**Agent behavior:**

- Create or update an architecture decision record for meaningful architectural changes.
- Include context, decision, alternatives considered, consequences, and status.
- Keep decision records concise.
- Link implementation work to the decision when possible.
- Update superseded decisions instead of silently contradicting them.

**Acceptance criteria:**

- Future maintainers can understand why the decision was made.
- Tradeoffs are explicit.
- Rejected alternatives are named.
- Consequences are acknowledged.

**Bad smells:**

- Architecture only explained in chat or PR comments.
- Repeated debates because prior decisions are invisible.
- Decision records that state the outcome but not the reasoning.
- Stale decisions that conflict with current code.

---

## RG-17 — Maintain repository-level AI-agent instructions

**Core rule:** Agentic coding works better when repository-specific rules, commands, architecture, and validation expectations are explicit.

**Agent behavior:**

- Prefer existing agent instruction files over guessing.
- Add or update an agent instruction file when the project lacks one and agent usage is expected.
- Include build, test, lint, run, and validation commands.
- Include architecture boundaries, risky files, generated files, and style expectations.
- Use path-specific instructions for meaningfully different subdomains.

**Acceptance criteria:**

- A new agent can make a safe first change without reverse-engineering everything.
- Validation commands are discoverable.
- Generated or vendored code is clearly marked.
- Project-specific constraints override generic habits.

**Bad smells:**

- Agents repeatedly guessing setup commands.
- Hidden tribal knowledge.
- AI-generated diffs that violate local architecture.
- No clear distinction between source, generated files, fixtures, and build output.

---

# User-Mandated Design Principles

These rules were requested directly and do not require the three-source research gate.

## UM-01 — Separation of concerns

**Core rule:** Keep different responsibilities separate unless combining them clearly reduces complexity.

**Agent behavior:**

- Separate domain logic from I/O, UI, persistence, transport, and infrastructure.
- Keep policy decisions separate from low-level mechanics.
- Avoid mixing validation, transformation, side effects, and presentation in one block when they change for different reasons.

**Good outcome:** A change in storage should not require rewriting business logic. A UI change should not require rewriting the domain model.

---

## UM-02 — Logical folder structure

**Core rule:** Folder layout should reflect the way the system is understood and changed.

**Agent behavior:**

- Follow the existing project layout unless it is clearly broken.
- Group related code by feature, domain, layer, or runtime boundary according to project style.
- Keep tests, fixtures, docs, generated files, and source files easy to locate.
- Avoid creating new top-level concepts casually.

**Good outcome:** A maintainer can guess where a change belongs before searching.

---

## UM-03 — Do not overuse monolithic files

**Core rule:** A file should not carry too many unrelated responsibilities.

**Agent behavior:**

- Split files when independent concepts are forced to share space.
- Extract cohesive units with clear names.
- Prefer boundaries that match responsibilities, not arbitrary line counts.
- Keep public entry points stable when splitting internals.

**Good outcome:** Files are small enough to understand, test, and modify without scrolling through unrelated systems.

---

## UM-04 — Do not overuse separate files

**Core rule:** Too many tiny files can be as harmful as one giant file.

**Agent behavior:**

- Do not split code merely because it is possible.
- Keep tightly coupled logic together when separation would obscure the flow.
- Avoid one-class, one-function, or one-constant files unless the project style or domain warrants it.
- Prefer discoverability over artificial purity.

**Good outcome:** The reader does not need to bounce across many files to understand one simple behavior.

---

## UM-05 — Avoid backwards-compatibility hacks and one-off solutions

**Core rule:** Do not pile permanent complexity onto the codebase to preserve accidental behavior or satisfy a single awkward case.

**Agent behavior:**

- Prefer intentional migrations over hidden compatibility shims.
- Time-box compatibility layers and document removal plans.
- Avoid special-case branches that bypass the system’s normal model.
- Fix the underlying abstraction when several one-offs point to the same design gap.

**Good outcome:** Compatibility is deliberate, documented, and temporary where possible.

---

## UM-06 — Use unified solutions that are broadly applicable

**Core rule:** Prefer one coherent mechanism over many local exceptions.

**Agent behavior:**

- Solve repeated problems at the correct shared layer.
- Consolidate duplicate patterns when doing so reduces complexity.
- Make shared solutions configurable only where real variation exists.
- Avoid creating parallel systems that do almost the same thing.

**Good outcome:** New cases can use the standard path instead of adding another bespoke branch.

---

# Agentic Change Checklist

Use this before presenting a completed change.

## Before editing

- [ ] I inspected the relevant files.
- [ ] I identified the existing architecture and style.
- [ ] I found the smallest coherent change.
- [ ] I know which tests or checks should run.
- [ ] I checked for repository-specific agent instructions.

## During editing

- [ ] The diff stays focused on the task.
- [ ] Names and style match the project.
- [ ] New behavior has tests or a stated validation path.
- [ ] Public contracts, docs, or config examples are updated if needed.
- [ ] No secrets, local paths, or credentials were introduced.
- [ ] No unrelated refactors were mixed in.

## After editing

- [ ] I reviewed the diff myself.
- [ ] I ran targeted validation.
- [ ] I ran broader validation when practical.
- [ ] I investigated failures or disclosed blockers.
- [ ] I summarized what changed, why, and how it was validated.

---

# Review Rubric for Humans and Agents

Score each area as **pass**, **needs work**, or **not applicable**.

| Area | Review question |
|---|---|
| Purpose | Does the change have one clear goal? |
| Correctness | Does it satisfy the requested behavior? |
| Simplicity | Is this the simplest durable solution? |
| Scope | Are unrelated edits excluded? |
| Tests | Do tests cover the changed behavior and important edge cases? |
| Validation | Were the right commands run and reported? |
| Security | Are inputs, secrets, permissions, and errors handled safely? |
| Interfaces | Are public contracts preserved or deliberately changed? |
| Observability | Can important failures be diagnosed? |
| Documentation | Did relevant docs, examples, or ADRs change with the code? |
| Maintainability | Will a future maintainer understand this quickly? |
| Agent fit | Would another agent have enough context to continue safely? |

---

# Source Matrix

Each research-gated practice below is backed by at least three independent sources. The source links are included so a human or agent can audit the basis for each rule.

| ID | Practice | Source 1 | Source 2 | Source 3 |
|---|---|---|---|---|
| RG-01 | Small, focused changes | [Google Engineering Practices — Small CLs](https://google.github.io/eng-practices/review/developer/small-cls.html) | [Microsoft Engineering Playbook — Pull Requests](https://microsoft.github.io/code-with-engineering-playbook/code-reviews/pull-requests/) | [GitLab — Continuous integration best practices](https://about.gitlab.com/topics/ci-cd/continuous-integration-best-practices/) |
| RG-02 | Reviewable code before merge | [Google Engineering Practices — Code Review](https://google.github.io/eng-practices/review/) | [Google Engineering Practices — What to look for in a review](https://google.github.io/eng-practices/review/reviewer/looking-for.html) | [Atlassian — Code review best practices](https://www.atlassian.com/blog/loom/code-review-best-practices-2) |
| RG-03 | Deterministic validation | [OpenAI — Prompt guidance for coding agents](https://developers.openai.com/api/docs/guides/prompt-guidance) | [Microsoft — Unit testing best practices](https://learn.microsoft.com/en-us/dotnet/core/testing/unit-testing-best-practices) | [GitLab — Continuous integration best practices](https://about.gitlab.com/topics/ci-cd/continuous-integration-best-practices/) |
| RG-04 | Tests as behavior contracts | [Microsoft — Unit testing best practices](https://learn.microsoft.com/en-us/dotnet/core/testing/unit-testing-best-practices) | [Martin Fowler — Test Pyramid](https://martinfowler.com/bliki/TestPyramid.html) | [Microsoft Azure Well-Architected — Testing](https://learn.microsoft.com/en-us/azure/well-architected/operational-excellence/testing) |
| RG-05 | Simplicity over unnecessary generality | [Google Engineering Practices — What to look for in a review](https://google.github.io/eng-practices/review/reviewer/looking-for.html) | [Martin Fowler — YAGNI](https://martinfowler.com/bliki/Yagni.html) | [Microsoft — Code metrics and cyclomatic complexity](https://learn.microsoft.com/en-us/dotnet/fundamentals/code-analysis/quality-rules/ca1502) |
| RG-06 | Clear names and consistent style | [Google C++ Style Guide](https://google.github.io/styleguide/cppguide.html) | [Python PEP 8](https://peps.python.org/pep-0008/) | [Microsoft — C# coding conventions](https://learn.microsoft.com/en-us/dotnet/csharp/fundamentals/coding-style/coding-conventions) |
| RG-07 | Explicit contracts and stable interfaces | [OpenAPI Initiative](https://www.openapis.org/) | [Google Cloud API Design Guide](https://docs.cloud.google.com/apis/design) | [Microsoft Azure — API design](https://learn.microsoft.com/en-us/azure/architecture/best-practices/api-design) |
| RG-08 | Documentation as deliverable | [GitHub Docs — About READMEs](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-readmes) | [Google Developer Documentation Style Guide](https://developers.google.com/style) | [Microsoft Azure Well-Architected — Architecture decision records](https://learn.microsoft.com/en-us/azure/well-architected/architect-role/architecture-decision-record) |
| RG-09 | Version-control hygiene | [Git SCM — Git](https://git-scm.com/) | [Atlassian Community — Git best practices](https://community.atlassian.com/forums/Bitbucket-articles/Git-Best-Practices/ba-p/1628803) | [Tao et al. — Commit message literature review](https://arxiv.org/abs/2202.02974) |
| RG-10 | Dependency and supply-chain discipline | [OWASP Dependency-Check](https://owasp.org/www-project-dependency-check/) | [OWASP — Vulnerable Dependency Management Cheat Sheet](https://cheatsheetseries.owasp.org/cheatsheets/Vulnerable_Dependency_Management_Cheat_Sheet.html) | [GitHub Docs — Dependabot security updates](https://docs.github.com/en/code-security/concepts/supply-chain-security/dependabot-security-updates) |
| RG-11 | Config and secrets outside source | [The Twelve-Factor App — Config](https://12factor.net/config) | [OWASP — Secrets Management Cheat Sheet](https://cheatsheetseries.owasp.org/cheatsheets/Secrets_Management_Cheat_Sheet.html) | [GitHub Docs — Secret scanning](https://docs.github.com/code-security/secret-scanning/about-secret-scanning) |
| RG-12 | Secure-by-design defaults | [NIST SSDF SP 800-218](https://csrc.nist.gov/pubs/sp/800/218/final) | [OWASP Secure Coding Practices Checklist](https://owasp.org/www-project-secure-coding-practices-quick-reference-guide/stable-en/02-checklist/05-checklist) | [CISA — Secure by Design](https://www.cisa.gov/securebydesign) |
| RG-13 | Observability by design | [OpenTelemetry — Observability primer](https://opentelemetry.io/docs/concepts/observability-primer/) | [Google SRE Book — Monitoring Distributed Systems](https://sre.google/sre-book/monitoring-distributed-systems/) | [Microsoft Azure Well-Architected — Observability](https://learn.microsoft.com/en-us/azure/well-architected/operational-excellence/observability) |
| RG-14 | Safe refactoring | [Refactoring.com](https://refactoring.com/) | [Microsoft Visual Studio Docs — Refactoring](https://learn.microsoft.com/en-us/visualstudio/ide/refactoring-in-visual-studio) | [Martin Fowler — Refactoring, 2nd Edition](https://martinfowler.com/books/refactoring.html) |
| RG-15 | CI as quality gate | [GitLab — CI/CD](https://about.gitlab.com/topics/ci-cd/) | [Atlassian — Continuous integration](https://www.atlassian.com/continuous-delivery/continuous-integration) | [Atlassian — Continuous delivery and testing](https://www.atlassian.com/continuous-delivery) |
| RG-16 | Architecture decisions recorded | [Microsoft Azure Well-Architected — ADRs](https://learn.microsoft.com/en-us/azure/well-architected/architect-role/architecture-decision-record) | [Architecture Decision Record GitHub organization](https://github.com/architecture-decision-record/architecture-decision-record) | [Nogueira et al. — Architecture decision records in practice](https://arxiv.org/abs/2604.27333) |
| RG-17 | Repository-level AI-agent instructions | [OpenAI — Prompt guidance for coding agents](https://developers.openai.com/api/docs/guides/prompt-guidance) | [Anthropic — Claude Code best practices](https://code.claude.com/docs/en/best-practices) | [GitHub Docs — Repository custom instructions for Copilot](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/add-custom-instructions/add-repository-instructions) |

---

# Compact Agent Instruction Template

A repository can copy this section into `AGENTS.md`, `CLAUDE.md`, or equivalent.

```md
# Agent Instructions

## Mission
Make the smallest correct change that satisfies the task. Preserve existing architecture unless the task explicitly requires changing it.

## Before editing
- Read the relevant files first.
- Follow existing patterns.
- Identify tests and validation commands.
- Avoid unrelated refactors.

## Coding rules
- Keep changes small and focused.
- Preserve separation of concerns.
- Use clear names and project style.
- Do not introduce secrets or machine-specific paths.
- Prefer unified project-wide solutions over one-off patches.
- Avoid speculative abstractions.
- Update docs when behavior, setup, API, or architecture changes.

## Validation
Run targeted tests first, then broader checks when practical.
Report exact commands and results.
If a check cannot run, explain why and name the next best check.

## Final response
Include:
- What changed.
- Why it changed.
- Files touched.
- Tests/checks run.
- Risks or follow-up work.
```

---

# Practical Defaults for Agentic AI

Use these defaults when the repository does not specify otherwise.

- Prefer editing existing files over creating new files.
- Create a new file only when it introduces a distinct cohesive responsibility.
- Avoid large generated rewrites unless the task is explicitly a generation task.
- Never silently delete user work.
- Never ignore failing tests that are plausibly related to the change.
- Do not invent commands, dependencies, APIs, or configuration values.
- Mark uncertainty clearly.
- Preserve backwards compatibility only when it is a real supported contract, not accidental legacy behavior.
- If compatibility is required, isolate it, document it, and define the removal path when possible.
- Prefer project-wide mechanisms over local hacks.

---

# One-Page Summary

A good agentic code change is:

- **Small:** one coherent purpose.
- **Reviewable:** easy to understand and reason about.
- **Validated:** tests or checks prove the changed behavior.
- **Consistent:** follows existing architecture, names, and style.
- **Secure:** does not leak secrets or widen trust boundaries casually.
- **Observable:** important failures can be diagnosed.
- **Documented:** behavior and decisions are not trapped in chat history.
- **Unified:** avoids one-off hacks when a shared solution is warranted.
- **Balanced:** avoids both giant monolithic files and excessive file fragmentation.
