// Shared configuration for the browser build.
//
// XLIDE's container engine - the ZIP and CFB readers, the VBA project writer,
// the form designer storage - is dependency-free TypeScript that needs Node
// for exactly two things: compression, and reading bundled asset files. Each
// of those lives behind one leaf module with a desktop and a web twin, and
// this plugin is what swaps them.
//
// esbuild.js uses this for the web bundle; tests/webBundle.test.ts uses it to
// assert the engine still bundles for a browser at all, so a stray `import *
// as fs` in src/vba/ fails a test rather than the next browser release.

const fs = require("fs");
const path = require("path");

const ROOT = __dirname;

/**
 * Runtime text assets the browser bundle carries, because there is no
 * extension directory to read them from. Only the webview templates: the
 * test-host sources and the Access design blobs belong to features the web
 * build does not have.
 */
const BUNDLED_ASSET_DIRS = ["assets/webview"];

/** Reads those directories into the map src/webview/bundledAssets.ts declares. */
function collectBundledAssets() {
  const assets = {};
  for (const dir of BUNDLED_ASSET_DIRS) {
    const absolute = path.join(ROOT, dir);
    if (!fs.existsSync(absolute)) {
      continue;
    }
    for (const name of fs.readdirSync(absolute)) {
      const file = path.join(absolute, name);
      if (!fs.statSync(file).isFile()) {
        continue;
      }
      // Line endings normalized to LF, exactly as readExtensionTextAsset does,
      // so a webview renders the same however the repo was checked out.
      assets[`${dir}/${name}`] = fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n");
    }
  }
  return assets;
}

/**
 * Desktop leaf -> web twin. Both sides of each pair export the same API; see
 * the header comment in either file for why the split exists.
 */
const WEB_LEAVES = [
  {
    // node:zlib -> pure-TypeScript inflate, storing entries it cannot deflate.
    match: /containerCodecNode$/,
    replacement: "src/vba/containerCodecWeb.ts",
  },
  {
    // node:fs and node:crypto -> a cache the extension layer primes and drains.
    match: /hostPlatformNode$/,
    replacement: "src/vba/hostPlatformWeb.ts",
  },
  {
    // assets read from the extension directory -> assets carried in the bundle.
    match: /extensionAssets$/,
    replacement: "src/extensionAssetsWeb.ts",
  },
  {
    // node:path -> POSIX-only path logic; a web workspace has no other kind.
    match: /^(node:)?path$/,
    replacement: "src/util/webPath.ts",
  },
  {
    // node:crypto -> synchronous SHA-256/SHA-1, which SubtleCrypto is not.
    match: /^(node:)?crypto$/,
    replacement: "src/util/webCrypto.ts",
  },
  {
    // node:fs existence check -> workspace.fs, the only thing that can see a
    // virtual workspace's files.
    match: /fsNode$/,
    replacement: "src/util/fsWeb.ts",
  },
  {
    // Nothing -> loading the container through workspace.fs before the
    // synchronous engine runs, and writing back after. This is what makes the
    // engine reachable at all in a browser.
    match: /enginePrimingNode$/,
    replacement: "src/enginePrimingWeb.ts",
  },
  {
    // Launching Office -> refusing to. The write coordinator reaches this to
    // reopen a file in its host application, which a browser has none of.
    match: /officeHostLauncher$/,
    replacement: "src/officeHostLauncherWeb.ts",
  },
  {
    // `process.platform` -> 'web'. A browser has no `process` at all.
    match: /osPlatformNode$/,
    replacement: "src/util/osPlatformWeb.ts",
  },
  {
    // Settings, exports and backups -> workspace.fs, which is the only thing
    // that can see a virtual workspace's files.
    match: /workspaceFilesNode$/,
    replacement: "src/util/workspaceFilesWeb.ts",
  },
  {
    // The PowerShell spawner -> one that refuses. A browser has no shell.
    match: /powershellNode$/,
    replacement: "src/util/powershellWeb.ts",
  },
  {
    // The features that need a shell, a local Office install, a git binary or
    // a worker thread -> implementations that decline. This is the swap that
    // keeps child_process, worker_threads and the test host out of the web
    // bundle entirely.
    match: /platformFeaturesNode$/,
    replacement: "src/platformFeaturesWeb.ts",
  },
];

/**
 * Modules esbuild injects into the browser bundle, which rewrites free
 * identifiers they export. `Buffer` is a Node global the engine uses in 418
 * places; src/util/webBuffer.ts supplies it.
 */
const WEB_INJECT = ["src/util/webBufferInject.ts"];

/** Node builtins that must never reach the browser bundle. */
const FORBIDDEN_BUILTINS = [
  "zlib",
  "fs",
  "path",
  "os",
  "net",
  "http",
  "child_process",
  "worker_threads",
  "crypto",
  "util",
];

/** @returns {import('esbuild').Plugin} */
function webLeafSwap() {
  return {
    name: "xlide-web-leaf-swap",
    setup(build) {
      for (const leaf of WEB_LEAVES) {
        build.onResolve({ filter: leaf.match }, () => ({
          path: path.join(ROOT, leaf.replacement),
        }));
      }

      // The webview templates are read from disk on a desktop and carried in
      // the bundle here, so the module that declares them empty is replaced
      // with one that holds them.
      build.onLoad({ filter: /[\\/]webview[\\/]bundledAssets\.ts$/ }, () => ({
        contents:
          "export const bundledTextAssets: Record<string, string> = " +
          `${JSON.stringify(collectBundledAssets(), null, 1)};\n`,
        loader: "ts",
      }));
    },
  };
}

/** Absolute paths for esbuild's `inject`. */
function webInject() {
  return WEB_INJECT.map((relative) => path.join(ROOT, relative));
}

module.exports = {
  webLeafSwap,
  webInject,
  WEB_LEAVES,
  WEB_INJECT,
  FORBIDDEN_BUILTINS,
  ROOT,
};
