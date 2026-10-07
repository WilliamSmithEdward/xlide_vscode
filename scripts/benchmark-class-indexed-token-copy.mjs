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
const classes = [{ name: "Class1", moduleName: "Class1", kind: "class", members: [{ name: "M", moduleName: "Class1", kind: "property", returns: "Object", knownValue: "nothing" }] }];
for (const count of [1, 100, 600]) {
  const terms = Array(count).fill("actor.M(1)"), lines = [];
  for (let i = 0; i < count; i += 30) lines.push(terms.slice(i, i + 30).join("; "));
  const source = ["Option Explicit", "Sub Go()", "Dim actor As New Class1", "Debug.Print " + lines.join("; _\n"), "End Sub", ""].join("\n");
  const mod = api.parseModule(source), symbols = api.buildModuleSymbols("M", "standard", source, { parsedModule: mod });
  const runRule = () => {
    const out = [];
    api.checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, void 0, (...args) => out.push(args));
    return out;
  };
  let from = 0;
  const expectedRule = Array.from({ length: count }, () => {
    const start = source.indexOf("actor.M(1)", from);
    from = start + 10;
    return ["objectVariableNotSet", "'actor.M(1)' is Nothing here: nothing in Class1 sets M. This will raise Run-time error '91': Object variable or With block variable not set.", { start, end: from }];
  });
  const original = Array.prototype[Symbol.iterator];
  let tokenReferences = 0;
  Array.prototype[Symbol.iterator] = function() {
    if (this.length && typeof this[0]?.rawText === "string" && typeof this[0]?.kind === "string") tokenReferences += this.length;
    return original.call(this);
  };
  let actual;
  try {
    actual = runRule();
  } finally {
    Array.prototype[Symbol.iterator] = original;
  }
  assert.deepEqual(actual, expectedRule);
  if (!baseline) assert.ok(tokenReferences <= count * 10 + 20);
  workCounts.push({ count, tokenReferences, physicalLines: lines.length, continuations: lines.length - 1 });
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
      const begin = performance.now(), out = scope === "rule" ? runRule() : runModule(), elapsed = performance.now() - begin;
      assert.deepEqual(out, scope === "rule" ? expectedRule : expectedModule);
      if (round >= 0) samples.push(elapsed);
    }
    samples.sort((a, b) => a - b);
    rows.push({ count, scope, medianMs: +samples[4].toFixed(5) });
  }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: "Warmed AST/symbols outside direct-rule timing; full-module calls include analyzer setup and no-error assertions. Independent exact rule messages/spans and N object-variable-not-set full diagnostics. No editor/cold/heap claim." }, null, 2));
