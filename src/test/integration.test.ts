// Entry point of the integration suite: esbuild bundles this file (and what
// it imports) into out/test/integration.test.js, which .vscode-test.mjs hands
// to mocha inside a real VS Code. The suites call the same provider commands
// the editor calls, so a passing run means the feature works in the product,
// not only that its internals were exercised.

import './formatting.test';
import './deadCode.test';
import './gitCompare.test';
import './officeIntegration.test';
import './crossModuleDiagnostics.test';
import './canonicalCasing.test';
import './agentReview.test';
import './outsideChanges.test';
import './saveConflicts.test';
import './renameModule.test';
import './vbeRename.test';
import './treeFollow.test';
import './deleteModule.test';
import './docComments.test';
import './projectReferences.test';
import './shapes.test';
