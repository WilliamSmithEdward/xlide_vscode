#!/usr/bin/env node
// Regenerates src/analyzer/runtime/vbaRuntimeDocs.ts: the one-line reference
// summary for every built-in VBA function and statement the analyzer knows.
//
// Unlike the other reference generators, the source is not a repo-local type
// library dump - a type library carries no prose - but the published language
// reference, read from its own markdown source at MicrosoftDocs/VBA-Docs. The
// summary is the first prose paragraph of a name's page, taken verbatim.
//
// Usage: node scripts/generate-vba-runtime-docs.mjs   (requires network)

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outputPath = path.join(root, 'src', 'analyzer', 'host', '..', 'runtime', 'vbaRuntimeDocs.ts');
const BASE = 'https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/Language/Reference/User-Interface-Help';

console.log(
    'This generator reads the published reference over the network.\n' +
    `Source: ${BASE}\n` +
    `Target: ${path.relative(root, outputPath)}\n` +
    'It is intentionally manual: the descriptions change only when Microsoft\n' +
    'edits the reference, and the generated file is committed.',
);
