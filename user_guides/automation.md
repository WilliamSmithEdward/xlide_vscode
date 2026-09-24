# XLIDE Automation And CI Guide

XLIDE exposes file analysis and VBA test execution to AI agents through VS
Code language-model tools. Use these tools when an automated workflow needs to
inspect, edit, analyze, and test a macro-enabled Office file - Excel, Word,
PowerPoint, or Access - without driving the XLIDE panels by hand.

## Recommended Agent Flow

1. Discover the target file with `xlide_listProjects` (it lists every macro
   container and every VB6 `.vbp` project) or confirm structure with
   `xlide_getProjectInfo`.
2. Read the file's VBA with `xlide_readModule`.
3. Write changes with `xlide_writeModule` or the other module tools. A write
   to an Access file takes effect when Access next opens the database and
   recompiles. Each chat-driven write opens a before/after diff for the user,
   and the XLIDE tree badges the module (` ● agent edit`) with inline Review,
   Keep and Revert actions until they decide - no notifications, and reverting
   a module the write created removes it. The diff closes by itself once the
   change it shows is gone: the agent deletes a scratch module it created,
   puts a module back as it was, or the user reverts it. Only writes through
   XLIDE's tools are reviewed this way; XLIDE cannot tell an agent's edit
   typed into an open editor from the user's own. `xlide.agent.showWriteDiffs`
   turns the review off.
4. Run `xlide_analyzeProject` and treat an empty `problems` array as analysis
   pass.
5. Run `xlide_runVbaTests` to execute discovered `@xlide-test` procedures
   through the production read-only test host of the file's own application
   (Excel, Word, PowerPoint, or Access).
6. Before a commit, call `xlide_gitChanges` to see what changed inside the
   file since HEAD (or any revision): one unified diff per module, plus which
   modules were added or removed. `git diff` cannot say this about a binary file,
   so this is what an agent reviews changes or writes a commit message from.

## Agents On The MCP Server

An agent that is not Copilot can edit through the
[xlide-mcp](https://github.com/WilliamSmithEdward/xlide_mcp) server instead.
From xlide_mcp 1.1.0 the server tells XLIDE what each write changed, and
XLIDE shows it the way it shows its own tools' writes: open modules reload,
and each module the server wrote gets the before/after diff and the Review,
Keep and Revert actions in the tree. A module the server renames keeps its
review under the new name, and one it deletes ends it. This applies to a file
the window shows, in its tree or with a module of it open.

The toggle is **Mirror MCP server edits**, in the Agent Instructions dialog
of the XLIDE sidebar, and the setting is `xlide.agent.mirrorMcpEdits`. It is
on by default. While it is on, the window listens on a local port
(127.0.0.1) behind a random token, which it keeps in a file only you can
read. Turning it off closes the port. The diff and the Keep and Revert
actions also follow `xlide.agent.showWriteDiffs`.

`xlide_runVbaTests` supports `moduleName`, `procedureName`, `testIds`,
`includeTags`, `excludeTags`, and `failFast` so an agent can start narrow while
editing and finish with a full run.

## Test Artifacts

Agent-driven test runs write the same artifact surface as the Tests GUI:

```text
tests/
  file_name_yyyy-mm-dd_hhmmss/
    summary.json
    host-trace.json
    output.log
  status_for_ci.json
```

The returned JSON includes `{ ok, summary, artifacts, report }`.
`artifacts.ciStatus` is the same payload written to `status_for_ci.json`.
Downstream CI should prefer `status_for_ci.json` for the latest run and
`summary.json` when it needs the full result detail.

If setup is incomplete, the tool returns `blocked: true` with `reason:
"test-support"` or `reason: "office-com"` instead of attempting to start the
file's application.

## Analysis Contract

`xlide_analyzeProject` uses the same analyzer that powers editor diagnostics.
Problems include module, line, column, severity, code, and message. Test
directives such as `expected-error` suppress only the intentionally expected
deterministic runtime diagnostic path inside that test procedure; unrelated
syntax, compile, style, or structural diagnostics remain visible.
