import { build } from "esbuild";
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { tmpdir, cpus } from "node:os";
import { createRequire } from "node:module";
import { performance } from "node:perf_hooks";
import assert from "node:assert/strict";
const baseline = process.argv.find((a) => a.startsWith("--baseline="))?.slice(11), scratch = mkdtempSync(join(tmpdir(), "xlide-class-index-")), file = join(scratch, "api.cjs");
let api;
try {
  const plugins = baseline ? [{ name: "baseline", setup(b2) {
    b2.onLoad({ filter: /[\\/]rules[\\/]classInstanceValues\.ts$/ }, (a) => ({ contents: execFileSync("git", ["show", baseline + ":src/analyzer/diagnostics/rules/classInstanceValues.ts"], { encoding: "utf8" }), loader: "ts", resolveDir: dirname(a.path) }));
  } }] : [];
  const b = await build({ plugins, stdin: { contents: "export {checkClassInstanceValues} from './src/analyzer/diagnostics/rules/classInstanceValues';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: "ts" }, bundle: true, platform: "node", format: "cjs", write: false });
  writeFileSync(file, b.outputFiles[0].contents);
  api = createRequire(import.meta.url)(file);
} finally {
  try {
    unlinkSync(file);
  } finally {
    rmdirSync(scratch);
  }
}
const rows = [], workCounts = [];
for (const kind of ["repeated", "distinct"]) for (const count of [1, 100, 1e3]) {
  let nameReads = 0;
  const classes = [{ name: "Class1", moduleName: "Class1", kind: "class", members: Array.from({ length: count }, (_, i) => ({ get name() {
    nameReads++;
    return "M" + i;
  }, moduleName: "Class1", kind: "property", returns: "Object", knownValue: "nothing" })) }];
  const source = [
    "Option Explicit",
    "Sub Go()",
    "Dim actor As New Class1",
    ...Array.from({ length: count }, (_, i) => "Debug.Print actor.M" + (kind === "distinct" ? i : count - 1) + ".Count"),
    "End Sub",
    ""
  ].join("\n");
  const mod = api.parseModule(source), symbols = api.buildModuleSymbols("M", "standard", source, { parsedModule: mod });
  const runRule = () => {
    const out = [];
    api.checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, void 0, (...args) => out.push(args));
    return out;
  };
  const expectedRule = Array.from({ length: count }, (_, i) => {
    const label = "actor.M" + (kind === "distinct" ? i : count - 1), line = "Debug.Print " + label + ".Count";
    const start = source.indexOf(line, source.indexOf("Dim actor")) + 12;
    const actualStart = kind === "distinct" ? start : source.indexOf(line) + 12 + i * (line.length + 1);
    return ["objectVariableNotSet", "'" + label + "' is Nothing here: nothing in Class1 sets M" + (kind === "distinct" ? i : count - 1) + ". This will raise Run-time error '91': Object variable or With block variable not set.", { start: actualStart, end: actualStart + label.length }];
  });
  nameReads = 0;
  assert.deepEqual(runRule(), expectedRule);
  workCounts.push({ kind, count, nameReads });
  if (!baseline) assert.ok(nameReads <= count * 2);
  const runModule = () => {
    const failures = [];
    const out = api.analyzeModule(source, { projectClassMembers: classes, onInternalError: (e) => failures.push(String(e)) });
    assert.deepEqual(failures, []);
    return out;
  };
  const expectedModule = runModule();
  assert.equal(expectedModule.length, count);
  assert.ok(expectedModule.every((d) => d.code === "object-variable-not-set"));
  for (const scope of ["rule", "complete-module-diagnostics"]) {
    const samples = [];
    for (let round = -3; round < 9; round++) {
      const begin = performance.now(), actual = scope === "rule" ? runRule() : runModule(), elapsed = performance.now() - begin;
      assert.deepEqual(actual, scope === "rule" ? expectedRule : expectedModule);
      if (round >= 0) samples.push(elapsed);
    }
    samples.sort((a, b) => a - b);
    rows.push({ kind, count, scope, medianMs: +samples[4].toFixed(5) });
  }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: "Warm AST/symbols outside direct-rule timing; complete-module calls include analyzer setup and no-error assertions. Independent exact direct-rule messages/spans and whole-module diagnostic count/code controls. No editor/cold/heap claim." }, null, 2));
