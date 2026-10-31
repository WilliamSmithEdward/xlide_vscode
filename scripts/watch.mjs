#!/usr/bin/env node
// Run the esbuild and tsc watchers together for the default build task
// (.vscode/tasks.json). Both write to this terminal, and when either stops
// the other is stopped too, so the task never keeps a half-working watch.
//
// This replaces npm-run-all, whose minimatch 3 carries GHSA-g3hv-75wg-m249
// with no fixed release on that line.
import { spawn } from 'node:child_process';

const children = ['watch:esbuild', 'watch:tsc'].map((script) =>
  spawn('npm', ['run', script], { stdio: 'inherit', shell: true }),
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
