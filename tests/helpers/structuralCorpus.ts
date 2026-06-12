// Repository-wide VBA sample collection for the audit #74 engine-comparison
// harnesses: every ```vba block in syntax_corpus/*.md, every oracle case
// source, and every .bas/.cls/.frm fixture under excel_test_workbook/.

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

export interface CorpusSample {
    /** Stable identifier for reporting, e.g. "file.md#3" or an oracle case id. */
    id: string;
    source: string;
}

const repoRoot = process.cwd();
const corpusRoot = join(repoRoot, 'syntax_corpus');
const fixtureRoot = join(repoRoot, 'excel_test_workbook');

function markdownVbaBlocks(): CorpusSample[] {
    const samples: CorpusSample[] = [];
    for (const entry of readdirSync(corpusRoot)) {
        if (!entry.endsWith('.md')) {
            continue;
        }
        const text = readFileSync(join(corpusRoot, entry), 'utf8');
        const fence = /```vba\r?\n([\s\S]*?)```/g;
        let match: RegExpExecArray | null;
        let index = 0;
        while ((match = fence.exec(text)) !== null) {
            samples.push({ id: `${entry}#${index}`, source: match[1] });
            index++;
        }
    }
    return samples;
}

function oracleCaseSources(): CorpusSample[] {
    const oracle = JSON.parse(
        readFileSync(join(corpusRoot, 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
    ) as {
        cases: {
            id: string;
            source?: string;
            modules?: { name: string; source: string }[];
        }[];
    };
    const samples: CorpusSample[] = [];
    for (const c of oracle.cases) {
        if (typeof c.source === 'string') {
            samples.push({ id: `oracle:${c.id}`, source: c.source });
        }
        for (const mod of c.modules ?? []) {
            samples.push({ id: `oracle:${c.id}/${mod.name}`, source: mod.source });
        }
    }
    return samples;
}

function workbookFixtureModules(): CorpusSample[] {
    const samples: CorpusSample[] = [];
    // excel_test_workbook/ is a local-only fixture (gitignored), absent on CI
    // runners; the comparison harnesses must not depend on its presence.
    if (!existsSync(fixtureRoot)) {
        return samples;
    }
    const walk = (dir: string): void => {
        for (const entry of readdirSync(dir, { withFileTypes: true })) {
            const path = join(dir, entry.name);
            if (entry.isDirectory()) {
                walk(path);
                continue;
            }
            if (/\.(bas|cls|frm)$/i.test(entry.name) && statSync(path).size < 512 * 1024) {
                samples.push({
                    id: path.slice(repoRoot.length + 1).replace(/\\/g, '/'),
                    source: readFileSync(path, 'utf8'),
                });
            }
        }
    };
    walk(fixtureRoot);
    return samples;
}

/** Every VBA sample in the repository, used to diff engine behavior. */
export function allStructuralComparisonSamples(): CorpusSample[] {
    return [...markdownVbaBlocks(), ...oracleCaseSources(), ...workbookFixtureModules()];
}
