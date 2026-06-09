# General Development Strategy Guide

This guide captures a practical development strategy for building software that stays coherent as it grows. It is intended to apply broadly across programming projects, products, tools, libraries, services, internal platforms, data systems, and automation systems.

The core idea is simple: prefer shared rules, clear ownership, and repeatable feedback loops over scattered fixes. A good system should become easier to extend over time, not harder.

## 1. Start With The Broader Concern

When a bug, request, or example appears, treat it as a signal rather than an isolated task.

Ask:

- What larger rule does this example reveal?
- Where else should the same behavior apply?
- Is this a symptom of duplicated logic, unclear ownership, or an incomplete model?
- Can this be solved once in a shared place?

Individual examples are useful because they make the problem concrete. The implementation should usually solve the broader concern behind the example, not just the exact input that exposed it.

## 2. Prefer Unified Business Rules

The same rule should behave the same way wherever the user or developer encounters it.

If a rule affects validation, formatting, completion, navigation, testing, UI state, API behavior, or generated output, avoid implementing separate interpretations in each surface. Define the rule once, expose it through a shared helper, and let every surface consume the same source of truth.

This avoids a common failure mode: one workflow appears fixed while the same underlying issue remains broken somewhere else.

## 3. Eliminate Secondary Pipelines

Multiple ways of doing the same thing are expensive. They create inconsistent behavior, duplicated tests, unclear ownership, and long-term maintenance drag.

When two code paths serve the same purpose, consolidate them unless there is a clear reason not to. A separate path should have an explicit contract, not merely exist because it was convenient at the time.

Good consolidation targets include:

- Formatting and normalization rules
- Validation and diagnostic rules
- Search, suggestion, or lookup logic
- Metadata loading and precedence
- Rename, reference, routing, and navigation behavior
- Serialization, parsing, or transformation logic

When adding new functionality, look for old paths that should be retired at the same time.

## 4. Make Confidence Explicit

Not all knowledge is equal. A system should distinguish between facts it knows, facts it infers, and facts it cannot safely determine.

Use different behavior for different confidence levels:

- Known: enforce strongly and surface precise diagnostics.
- Inferred: provide helpful suggestions, but avoid overclaiming.
- Unknown: degrade gracefully and avoid false certainty.

This is especially important for type systems, static analysis, user-facing diagnostics, API contracts, data validation, and generated recommendations.

## 5. Prefer Broad Helpers Over One-Off Fixes

Small targeted fixes are sometimes necessary, but a pattern of one-off fixes usually means the system lacks a central abstraction.

Before patching a specific case, ask whether the project needs a reusable helper such as:

- `isValidCallExpression`
- `resolveSymbol`
- `normalizeIdentifier`
- `getCanonicalName`
- `formatLineOnExit`
- `findOwningProject`
- `classifyToken`
- `buildCompletionItem`

The exact helper names will vary by system. The principle is stable: encode the rule once, give it a clear name, and reuse it.

## 6. Build For Everyday Workflow First

High-value development work often comes from improving the small interactions that happen constantly.

Prioritize behavior that affects daily flow:

- Editing and formatting
- Fast feedback
- Clear diagnostics
- Navigation
- Rename and refactor support
- Search and references
- Suggestions and assisted input
- Error recovery
- Predictable project structure

These features compound because they reduce friction every time someone works in the codebase or product.

## 7. Treat Formatting As A Boundary Event

Formatting should happen at predictable boundaries, not only through explicit commands.

Useful boundaries include:

- Pressing Enter
- Moving away from a line
- Saving a file
- Accepting a completion
- Completing a refactor
- Generating code
- Applying a quick fix

The same formatting rules should apply across these boundaries whenever possible. A user should not have to remember which action triggers the correct version of the code.

## 8. Make Automation Syntax-Aware

Automation should understand enough of the surrounding syntax to avoid creating invalid or surprising code.

Examples:

- Assisted input should know whether it is inserting a declaration, expression, statement, type, value, operation, or keyword.
- A formatter should know whether it is inside a string, comment, block, or nested expression.
- A refactor should know whether it is renaming a local symbol, class, file, module, package, route, or external dependency.
- A block generator should know whether a closing block already exists.

The more context-sensitive the feature, the more important it is to route through shared parsing, classification, or symbol-resolution logic.

## 9. Diagnostics Should Be Precise And Actionable

Diagnostics should explain the actual rule that was violated. They should avoid vague language, false positives, and noisy warnings that train users to ignore them.

A useful diagnostic usually answers:

- What is wrong?
- Why is it wrong in this context?
- Where is the relevant code?
- What change would fix it?

When possible, diagnostics should be paired with quick fixes, documentation, or examples.

## 10. Test The Rule, Not Just The Example

When a bug report gives one example, add tests for the general rule.

A strong test set includes:

- The reported case
- Similar valid cases
- Similar invalid cases
- Cross-file or cross-module behavior when relevant
- Boundary cases
- Regression cases for prior implementation paths

Tests should confirm that all surfaces using the shared rule behave consistently.

## 11. Use Real-World Oracles When Available

When implementing behavior that must match an external system, language, runtime, protocol, file format, or platform, use that system as an oracle whenever possible.

Examples:

- Compare parser behavior against the real compiler.
- Validate generated files with the real application.
- Check API behavior against official documentation or live contract tests.
- Confirm platform-specific behavior with a minimal executable example.

Use the oracle to establish facts, then encode those facts as automated tests.

## 12. Keep Projects And Contexts Isolated

When a workspace, platform, or repository can contain multiple projects, packages, services, documents, tenants, environments, or deployments, keep their data and behavior isolated unless cross-context behavior is explicitly supported.

Shared tooling should always know which context owns the current file, symbol, configuration, dependency, state, or output.

This prevents subtle bugs where data from one project leaks into another.

## 13. Define Metadata Precedence

Most mature systems combine information from multiple sources: source code, generated metadata, configuration files, external APIs, package manifests, schemas, runtime inspection, and defaults.

Define precedence clearly:

- Which source wins when sources disagree?
- Which sources are authoritative?
- Which sources are fallback-only?
- Which sources are context-local?
- Which sources are global?

Unclear precedence creates inconsistent behavior and makes debugging difficult.

## 14. Prefer Coherent Batches

Small changes are easy to review, but very narrow changes can leave the system inconsistent. When work touches a shared rule, a larger coherent batch is often better.

A good batch includes:

- The shared rule or abstraction
- The main behavior change
- Updates across affected workflows or surfaces
- Tests for the rule
- Removal or retirement of redundant paths
- Roadmap or documentation updates when relevant

The goal is not to make changes large for their own sake. The goal is to leave the system more coherent than it was.

## 15. Keep The Roadmap Honest

A roadmap should be a living contract, not a static wish list.

Update it when:

- A feature is completed
- A discovery changes priorities
- A task turns out to be less valuable than expected
- A new class of bugs reveals missing infrastructure
- A dependency, platform constraint, or core workflow changes the plan

Roadmaps are most useful when they track capabilities, risks, and sequencing rather than just tasks.

## 16. Refactor Conservatively But Deliberately

Refactoring is not separate from feature work. It is often the work required to make a feature correct.

Refactor when it:

- Removes duplicated behavior
- Centralizes a rule
- Makes future bugs less likely
- Clarifies ownership
- Simplifies testing
- Reduces the number of valid implementation paths

Avoid broad cleanup that is unrelated to the current goal. The best refactors are connected to a specific behavior improvement.

## 17. Optimize For Maintainer Understanding

Future maintainers should be able to understand why the code works, not just that it works.

Prefer:

- Clear names
- Small shared helpers
- Focused modules
- Tests that describe behavior
- Comments only where they explain non-obvious decisions
- Documentation that captures strategy, not every implementation detail

Code should make the common path obvious and the unusual path explicit.

## 18. Recommended Operating Loop

Use this loop for most meaningful development work:

1. Understand the user or system behavior.
2. Identify the broader rule behind the example.
3. Locate all surfaces affected by the rule.
4. Find existing helpers, pipelines, and tests.
5. Consolidate duplicated behavior when practical.
6. Implement the shared rule.
7. Apply it across relevant surfaces.
8. Add tests for the general rule.
9. Remove or retire redundant paths.
10. Update documentation or roadmap notes.
11. Verify with automated tests and, when useful, a real-world oracle.

## 19. Common Anti-Patterns

Watch for these signs that the system is drifting:

- The same behavior is implemented in several places.
- A fix only works in one UI or workflow.
- New tests cover only the exact reported string.
- Diagnostics guess when they should be silent.
- Formatting depends on which action triggered it.
- Metadata sources disagree without a defined winner.
- Refactors add abstraction without removing complexity.
- Roadmap items stay complete even after new evidence contradicts them.
- A feature works in one context but leaks into another.

These are not just code smells. They are signals that the system needs a clearer rule or a better boundary.

## 20. Decision Checklist

Before shipping a meaningful change, ask:

- Did we solve the broader rule, or only the visible example?
- Does every relevant surface use the same logic?
- Did we remove redundant paths where practical?
- Are confidence levels explicit?
- Are diagnostics precise and actionable?
- Are tests broad enough to catch nearby regressions?
- Does this change preserve context isolation?
- Is metadata precedence clear?
- Did the roadmap or documentation need an update?
- Would a future maintainer understand why this is the right behavior?

## Closing Principle

The most effective development strategy is to make the correct path the shared path.

When rules are centralized, surfaces are consistent, tests describe behavior, and redundant paths are retired, software becomes easier to extend. Each change then improves not only the current feature, but the system's ability to absorb the next one.
