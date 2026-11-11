import { defineConfig, type Plugin } from 'vitest/config';

// Each test file that imports the analyzer loads it afresh, and with it
// the generated type-library tables (src/analyzer/host/*ObjectModelData.ts
// and friends, about 12 MB). Vite gives each of them a source map larger
// than the code, which travels to the worker and is inlined into the module
// for every such test file. A generated file - one whose first line says
// so - gets no source map: a stack or a breakpoint inside one points at the
// transformed code rather than the file's own lines, and in exchange CI's
// Node 24 test run took 40% less time (November 2026). A run that starts
// with coverage on keeps every map, since coverage-v8 needs them to count
// the generated files' statements at all; coverage switched on later in a
// running vitest (an editor's coverage button) leaves the generated files
// uncounted.
function generatedTablesWithoutSourceMaps(): Plugin {
    let active = true;
    return {
        name: 'generated-tables-without-source-maps',
        enforce: 'pre',
        configureVitest({ vitest }) {
            active = !vitest.config.coverage.enabled;
        },
        transform(code, id) {
            if (!active || !/\.ts$/.test(id.split('?')[0]) || !/^\uFEFF?\/\/ (Generated (from|by)|GENERATED FILE)/.test(code)) {
                return null;
            }
            // An empty mapping tells Vite the module has no source map at all.
            return { code, map: { mappings: '' } };
        },
    };
}

export default defineConfig({
    plugins: [generatedTablesWithoutSourceMaps()],
    test: {
        environment: 'node',
        include: ['tests/**/*.test.ts'],
        coverage: {
            provider: 'v8',
            include: ['src/**'],
        },
    },
});
