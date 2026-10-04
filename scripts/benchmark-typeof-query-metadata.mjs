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
    b2.onLoad({ filter: /[\\/]rules[\\/]typeOfIs\.ts$/ }, (a) => ({ contents: execFileSync("git", ["show", baseline + ":src/analyzer/diagnostics/rules/typeOfIs.ts"], { encoding: "utf8" }), loader: "ts", resolveDir: dirname(a.path) }));
  } }] : [];
  const b = await build({ plugins, stdin: { contents: "export {checkTypeOfIsCompatibility} from './src/analyzer/diagnostics/rules/typeOfIs';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: "ts" }, bundle: true, platform: "node", format: "cjs", write: false });
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
  let names = 0, interfaces = 0;
  const classes = Array.from({ length: count + 1 }, (_, i) => ({ get name() {
    names++;
    return "Class" + (i + 1);
  }, moduleName: "Class" + (i + 1), kind: "class", members: [], get implements() {
    interfaces++;
    return [];
  } }));
  const targets = Array.from({ length: count }, (_, i) => "Class" + (kind === "distinct" ? i + 2 : 2));
  const source = ["Option Explicit", "Sub Go(ByVal actor As Class1)", ...targets.map((target) => "If TypeOf actor Is " + target + " Then\nDebug.Print 1\nEnd If"), "End Sub", ""].join("\n");
  const mod = api.parseModule(source), symbols = api.buildModuleSymbols("M", "standard", source, { parsedModule: mod });
  const runRule = () => {
    const out = [], visitor = api.checkTypeOfIsCompatibility(symbols, { projectClassMembers: classes }, (...args) => out.push(args));
    for (const procedure of mod.members) {
      if (procedure.kind !== "Procedure") continue;
      const visit = visitor(procedure);
      if (!visit) continue;
      for (const node of procedure.body) {
        if (node.kind === "IfBlock") {
          for (const branch of node.branches) if (branch.condition) visit(branch.condition);
        }
      }
    }
    return out;
  };
  let from = 0;
  const expectedRule = targets.map((target) => {
    const label = "TypeOf actor Is " + target, start = source.indexOf(label, from);
    from = start + label.length;
    return ["typeOfIsAlwaysFalse", "'TypeOf ... Is " + target + "' is always False: 'actor' is declared As Class1, which is never " + target + ".", { start, end: from }];
  });
  names = 0;
  interfaces = 0;
  assert.deepEqual(runRule(), expectedRule);
  workCounts.push({ kind, count, names, interfaces });
  if (!baseline) {
    assert.ok(names <= count * 4 + 12);
    assert.ok(interfaces <= count * 4 + 12);
  }
  // Work counters are deliberately excluded from timed metadata reads.
  for (let i = 0; i < classes.length; i++) {
    Object.defineProperty(classes[i], 'name', {value: 'Class' + (i + 1), writable: true, configurable: true, enumerable: true});
    Object.defineProperty(classes[i], 'implements', {value: [], writable: true, configurable: true, enumerable: true});
  }
  const runModule = () => {
    const errors = [];
    const diagnostics = api.analyzeModule(source, { projectClassMembers: classes, onInternalError: (e) => errors.push(String(e)) });
    assert.deepEqual(errors, []);
    return diagnostics;
  };
  const expectedModule = runModule();
  assert.equal(expectedModule.length, count);
  assert.ok(expectedModule.every((d) => d.code === "typeof-is-always-false"));
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
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: "Warm AST/symbols outside direct-rule timing. Complete-module calls include analyzer setup and error assertions. Exact independently expected rule messages/spans; full diagnostics independently N always-false warnings, with frozen complete-output equality outside timing. No editor/cold/heap claim." }, null, 2));
