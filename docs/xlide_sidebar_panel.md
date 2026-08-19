# XLIDE Sidebar Panel

Status: polished Activity Bar/sidebar WebviewView implemented; richer test and
health-probe sections remain roadmap work.

Purpose: define the XLIDE Activity Bar/sidebar experience for VBA
development in VS Code.

## Implemented Initial Slice

- `package.json` contributes a dedicated `xlide` Activity Bar container and the
  `xlide.sidebar` WebviewView.
- `src/xlideSidebarModel.ts` owns the testable sidebar section model.
- `src/xlideSidebar.ts` owns VS Code rendering and refresh behavior.
- The sidebar shows a compact welcome note pointing at the Explorer-hosted
  XLIDE workbook tree, then the target-workbook action group, core XLIDE
  actions, compact settings launchers, and support actions. Nothing is gated:
  the workbook engine runs in-process, so there is no dependency to install or
  probe before the sidebar is usable.
- Global/editor settings are opened through the dedicated XLIDE Global Settings
  GUI and still persist to VS Code machine/profile settings. Workbook-scoped
  settings continue to live in `<workbook>.xlide_settings.json` and are edited
  from workbook-facing GUIs rather than permanent sidebar rows.
- Dense inline global/machine settings rows have been replaced by one global
  settings launcher.

## Stability Agreement

The current sidebar information architecture is considered stable:

```text
Welcome
Setup
Workbook Actions
Settings
Support
```

Do not rework the sidebar layout, section order, or status/action split without
a specific product reason. Future sidebar work should be additive and tied to
real workflow functionality. The expected next Workbook Actions addition is a
Unit Tests button once the VBA test runner is implemented; it should use the
sidebar target workbook and follow the same disabled-state contract as the other
workbook-scoped actions.

## Goals

- Make XLIDE visible as a first-class Activity Bar extension.
- Put the most common workbook workflows in one polished sidebar.
- Surface setup health clearly before a developer tries to analyze, run, or test.
- Provide one unified XLIDE configuration surface while preserving native VS
  Code Settings for users who prefer it.
- Keep every status deterministic: pass, warn, fail, or unknown with a concrete
  next action.
- Avoid replacing command-palette workflows; the sidebar should make them easier
  to discover and run.
- Keep module-scoped actions and module-scoped information out of the sidebar.
  Module-targeted commands belong in the editor or workbook/module tree where
  the module target is explicit.
- Keep workbook/module navigation available in the VS Code Explorer. The XLIDE
  Activity Bar/sidebar is an additional command, health, test, analysis, and config
  surface, not a reason to remove the Explorer tree.

## Activity Bar Icon

The icon should follow standard VS Code Activity Bar aesthetics:

- monochrome SVG
- no embedded color
- readable at 24px
- simple line geometry
- theme-neutral
- works as an icon mask
- visually distinct from Explorer/Search/Source Control/Run/Test

Candidate concepts:

- workbook page with a small VBA/module glyph
- workbook page plus bracket/slide motif
- XLIDE monogram simplified into a single-line mark

The icon should live in a stable media asset path and be referenced through the
`viewsContainers.activitybar` contribution.

## Sidebar Structure

Recommended sections:

1. **Welcome**
   - Explain that workbook/module navigation lives in Explorer > XLIDE.

2. **Workbook Actions**
   - Target workbook picker for workbook-scoped sidebar actions
   - Analyze selected workbook
   - Import/export selected workbook modules
   - Open selected workbook in Excel
   - Open selected workbook read-only
   - Future workbook-scoped test-runner entry points

4. **Problems**
   - Error count
   - Warning count
   - Suppressed diagnostic count once suppression support exists
   - Last analysis run summary

5. **Tests**
   - Pass/fail/skip/xfail/xpass summary
   - Last test run
   - Rerun failed
   - Tag/filter entry points once the test runner exists

6. **Activity**
   - Recent XLIDE operations
   - Last export/import
   - Last analysis/test run
   - COM/reset warnings

7. **Configuration**
   - One action to open the dedicated XLIDE Global Settings GUI for VS
     Code/machine settings.
   - Do not render every setting as sidebar rows; detailed inspection/editing
     belongs in dedicated settings GUIs.
   - Profile/rule-set selector once profiles exist, likely inside the dedicated
     settings GUI rather than as permanent sidebar clutter.

8. **Support**
   - Copy diagnostics
   - Export support bundle

## Setup Health Checks

Each check must return a deterministic state:

```text
pass
warn
fail
unknown
```

Do not infer from filenames, workbook names, or likely intent. If XLIDE cannot
prove the state, show `unknown`.

Initial checks:

- XLIDE extension active.
- A workbook/project context is selected.
- Workbook file exists and is reachable.
- Source export mapping exists.
- Modules can be listed.
- Current module source can be read.
- Current module source can be written when editing is enabled.
- Workbook can be opened/reopened in XLIDE context.
- Excel COM is available on Windows for run/test workflows.
- Trust access to the VBA project object model is enabled when COM write/run
  workflows require it.
- Macro/security prerequisites are sufficient for run/test workflows.
- Analysis engine is available.
- Test runner is configured once implemented.
- Optional `.vbref.xml` or doc-comment metadata is loaded if present.

Dependency/status rows should be compact and actionable. They belong in the
sidebar when they answer "can I use XLIDE right now?" and offer a direct fix.
They should not expand into full setting forms; detailed configuration belongs
in the dedicated settings GUIs.

Examples:

```text
pass    Workbook linked
warn    Excel COM unavailable, run/test disabled on this machine
fail    Trust access to VBA project object model is disabled
unknown No workbook selected
```

## Primary Actions

Actions should be visible as buttons or tree-item commands, with tooltips and
disabled states where appropriate:

- `XLIDE: Analyze Workbook` against the sidebar target workbook
- `XLIDE: Analyze Workbook` from the workbook node context menu in the XLIDE
  workbook tree; this should route to the same workbook analysis engine and active
  workbook selection model as the sidebar action.
- `XLIDE: Run All VBA Tests` against the sidebar target workbook once the test
  runner exists
- `XLIDE: Export/Sync Modules` against the sidebar target workbook
- `XLIDE: Open Workbook` for the sidebar target workbook
- `XLIDE: Open Workbook Read Only` for the sidebar target workbook
- `XLIDE: Reopen Workbook`
- `XLIDE: Open Global Settings`
- `XLIDE: Open Logs`

Buttons should not silently run destructive operations. Sync/write actions need
clear status and should reuse existing safe workbook handling. Workbook-scoped
sidebar buttons must be disabled when no target workbook is selected instead of
falling back to an unrelated active editor.

Module-scoped commands such as analyze current module, export current module,
and run current test should remain available from editor context menus,
command-palette flows, or module tree context menus, not as permanent sidebar
buttons.

## Polish Details

- Use VS Code codicons where available.
- Use concise labels; put detail in tooltips.
- Use a polished WebviewView layout with stable section spacing, borders,
  dividers, compact status dots, and button-style action rows.
- Keep the sidebar short. Prefer one launch button for a detailed GUI over
  many always-visible setting rows.
- Keep dependency health visible when it affects whether the current workflow is
  usable; use green/yellow status icons and one direct action for yellow items.
- Keep status dots reserved for setup/dependency health; project/welcome and
  workbook-action rows should not show health-color indicators.
- Keep sections collapsible.
- Persist collapsed/expanded state locally.
- Show badges for errors, warnings, and failed tests.
- Prefer inline action buttons for common actions.
- Keep long-running actions cancelable where VS Code supports it.
- Show progress for analysis/test/sync operations.
- Link result rows to modules and exact source lines.
- Render workbook analysis results in a dedicated native-feeling GUI/panel with
  module grouping, severity filters, counts, suppressed-diagnostic visibility,
  and copy/export actions. The Output channel should remain a support log, not
  the main analysis-results surface.
- Keep all colors theme-driven.
- Do not use custom decorative UI that fights VS Code styling.

## Accessibility

- Every icon-only action needs a tooltip.
- Tree items should have readable labels and descriptions.
- Status should not rely on color alone.
- Keyboard navigation should work through standard VS Code tree/view behavior.

## Implementation Notes

- Use a VS Code `WebviewView` for the polished sidebar shell. Keep the model
  separate and dependency-free so the business shape remains unit-testable.
- Keep the sidebar state model separate from VS Code rendering so it can be unit
  tested.
- Reuse the existing workbook explorer provider where possible.
- Keep the workbook/module explorer contributed to the Explorer view. The
  dedicated XLIDE Activity Bar view can reference the same selected workbook and
  provide actions/status, but should not make the Explorer tree disappear.
- Treat the sidebar target workbook as the single workbook-scoped action target
  for analysis, import/export, open/reopen, and future test-runner actions. The
  picker should list workbooks discovered from the open VS Code workspace, clear
  invalid selections when files disappear, and refresh when workbook files or
  workbook sidecars change.
- Do not render module-scoped actions or module-scoped information in the
  sidebar. The sidebar can route to the Explorer tree or editor context for
  module workflows, but it should not own a second module target picker or infer
  module intent from the active editor.
- Keep tree right-click workbook actions and sidebar workbook actions in sync as
  parallel entry points. The tree supplies the clicked workbook node, the sidebar
  supplies the target picker value, and both must call the same command handlers
  and downstream GUI/business logic.
- Store selected-workbook UI state in VS Code workspace state, not global
  settings and not workbook sidecar settings.
- Setup checks should be pure functions where possible and explicit async probes
  where COM/workbook state is required.
- Do not require Excel COM for basic sidebar rendering.
- Do not run health probes that mutate the workbook.
- Dedicated settings GUIs should consume the same resolver used by production
  code so displayed values always match behavior. The resolver must preserve
  setting provenance and validate malformed VS Code settings or workbook
  sidecars deterministically. The sidebar should summarize status and launch
  those GUIs instead of duplicating their detailed controls.
- Workbook-facing GUIs must use the shared configuration resolver: global
  VS Code settings provide defaults, workbook-scoped sidecars provide overrides,
  and GUI actions tied to a workbook must not silently mutate global defaults.
- Keep analysis, import/export, diff, and future test GUI settings on the same
  scoping model. Avoid compatibility-only command/config paths that make one
  surface behave differently from another.

## Non-Goals

- The sidebar is not a replacement for inline diagnostics.
- The sidebar should not make heuristic guesses about project state.
- The sidebar should not run tests or macros automatically.
- The sidebar should not require COM to show analysis-only workflows.

## Definition Of Done

- XLIDE has a dedicated Activity Bar icon and sidebar container.
- Setup health is visible with deterministic pass/warn/fail/unknown states.
- Analyze, test, sync, workbook, and refresh actions are discoverable.
- Results and status link back to source where practical.
- The UI looks native in VS Code light, dark, and high-contrast themes.
- Workbook-facing settings preserve provenance, use global defaults only as
  fallback, and persist workbook-specific choices at workbook scope.
