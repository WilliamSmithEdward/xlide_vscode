const esbuild = require("esbuild");
const { webLeafSwap, webInject } = require("./webBuild.js");

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");

/** @type {import('esbuild').Plugin} */
const esbuildProblemMatcherPlugin = {
  name: "esbuild-problem-matcher",
  setup(build) {
    build.onStart(() => console.log("[watch] build started"));
    build.onEnd((result) => {
      result.errors.forEach(({ text, location }) => {
        console.error(`[ERROR] ${text}`);
        if (location) {
          console.error(`    ${location.file}:${location.line}:${location.column}`);
        }
      });
      console.log("[watch] build finished");
    });
  },
};

async function main() {
  // Only the extension host code is bundled. Runtime text assets -- the
  // webview templates under assets/webview/ and the test-host sources under
  // assets/testhost/ -- are NOT bundled or copied here: they are read from
  // disk at runtime via src/extensionAssets.ts relative to the extension
  // root, and they ship in the VSIX because .vscodeignore keeps assets/**.
  const ctx = await esbuild.context({
    // Two bundles: the extension host, and the analysis worker thread the host
    // spawns (out/analysisWorker.js) so full analysis passes run off-thread.
    // A development build adds the integration suite (out/test/), which the
    // VS Code test runner loads; a production build ships without it.
    entryPoints: [
      "src/extension.ts",
      "src/analysisWorker.ts",
      ...(production ? [] : ["src/test/integration.test.ts"]),
    ],
    bundle: true,
    format: "cjs",
    minify: production,
    sourcemap: !production,
    sourcesContent: false,
    platform: "node",
    outdir: "out",
    // Both come from the host at run time: `vscode` from the editor, `mocha`
    // from the integration test runner.
    external: ["vscode", "mocha"],
    logLevel: "silent",
    plugins: [esbuildProblemMatcherPlugin],
  });

  // The same src/extension.ts, built for the web extension host: no Node, so
  // the leaf modules webBuild.js names are swapped for their web twins and
  // `Buffer` is injected. There is deliberately no second entry point - one
  // activate() means the two platforms cannot drift apart. The analysis
  // worker has no counterpart here (a browser has no worker_threads; see
  // platformFeaturesWeb.ts), so only the extension itself is built.
  const webCtx = await esbuild.context({
    entryPoints: ["src/extension.ts"],
    bundle: true,
    format: "cjs",
    minify: production,
    sourcemap: !production,
    sourcesContent: false,
    platform: "browser",
    outfile: "out/web/extension.js",
    external: ["vscode"],
    inject: webInject(),
    plugins: [webLeafSwap(), esbuildProblemMatcherPlugin],
    logLevel: "silent",
  });

  if (watch) {
    await Promise.all([ctx.watch(), webCtx.watch()]);
  } else {
    await Promise.all([ctx.rebuild(), webCtx.rebuild()]);
    await Promise.all([ctx.dispose(), webCtx.dispose()]);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
