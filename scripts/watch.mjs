#!/usr/bin/env node
// Run the esbuild and tsc watchers together for the default build task
// (.vscode/tasks.json). Both write to this terminal, and when either stops
// the other is stopped too, so the task never keeps a half-working watch.
//
// This replaces npm-run-all, whose minimatch 3 carries GHSA-g3hv-75wg-m249
// with no fixed release on that line.
//
// Each watcher runs through npm's own CLI, which npm names in npm_execpath
// for every script it runs, so no shell starts it on any platform.
import { spawn } from 'node:child_process';

const npmCli = process.env.npm_execpath;
if (!npmCli) {
  console.error('Run this through npm: npm run watch');
  process.exit(1);
}

const children = ['watch:esbuild', 'watch:tsc'].map((script) =>
  spawn(process.execPath, [npmCli, 'run', script], { stdio: 'inherit' }),
);

let exiting = false;
for (const child of children) {
  child.on('exit', (code) => {
    if (exiting) {
      return;
    }
    exiting = true;
    for (const other of children) {
      if (other !== child) {
        other.kill();
      }
    }
    process.exitCode = code ?? 1;
  });
}
