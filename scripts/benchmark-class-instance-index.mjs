// node scripts/benchmark-class-instance-index.mjs [--baseline=COMMIT]
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
const classes = [{ name: "Class1", moduleName: "Class1", kind: "class", members: [] }], rows = [], workCounts = [];
for (const kind of ["scalar", "class"]) for (const count of [1, 100, 1e3]) {
  const source = ["Option Explicit", "Sub Go()", ...Array.from({ length: count }, (_, i) => "Dim actor" + i + " As " + (kind === "scalar" ? "Long" : "New Class1")), ...Array(count).fill("Debug.Print 1"), "End Sub", ""].join("\n"), mod = api.parseModule(source), symbols = api.buildModuleSymbols("M", "standard", source, { parsedModule: mod });
  const runRule = () => {
    const out = [];
    api.checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, void 0, (...args) => out.push(args));
    return out;
  };
  let rowVisits = 0;
  const originals = { filter: Array.prototype.filter, every: Array.prototype.every };
  for (const method of ["filter", "every"]) Array.prototype[method] = function(callback, thisArg) {
    return originals[method].call(this, (value, index, array) => {
      if (value && Array.isArray(value.toks) && value.span) rowVisits++;
      return callback.call(thisArg, value, index, array);
    });
  };
  let result;
  try {
    result = runRule();
  } finally {
    for (const method of ["filter", "every"]) Array.prototype[method] = originals[method];
  }
  assert.deepEqual(result, []);
  if (!baseline) assert.equal(rowVisits, 0);
  workCounts.push({ kind, count, rowVisits });
  const runModule = () => {
    const failures = [], diagnostics = api.analyzeModule(source, { projectClassMembers: classes, onInternalError: (e) => failures.push(String(e)) });
    assert.deepEqual(failures, []);
    assert.equal(diagnostics.length, count);
    assert.ok(diagnostics.every((d) => d.code === "unused-variable"));
    assert.deepEqual(diagnostics.map((d) => d.data.removeDeclaration.variableName), Array.from({ length: count }, (_, i) => "actor" + i));
    return diagnostics;
  };
  const expected = runModule();
  for (const scope of ["rule", "complete-module-diagnostics"]) {
    const samples = [];
    for (let round = -3; round < 9; round++) {
      const begin = performance.now(), actual = scope === "rule" ? runRule() : runModule(), elapsed = performance.now() - begin;
      assert.deepEqual(actual, scope === "rule" ? [] : expected);
      if (round >= 0) samples.push(elapsed);
    }
    samples.sort((a, b) => a - b);
    rows.push({ kind, count, scope, medianMs: +samples[4].toFixed(5), outputsChecked: true });
  }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: "Warmed module AST/symbols; pure rule result independently empty; whole-module results frozen with independent exact unused-name/code/count and no-error controls. Assertions inside full call affect timing. No cold/heap/editor claim." }, null, 2));
