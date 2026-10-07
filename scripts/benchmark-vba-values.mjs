// Run with node scripts/benchmark-vba-values.mjs; uses the pinned esbuild dependency.
import { build } from 'esbuild';
import { performance } from 'node:perf_hooks';

const bundle = await build({
  stdin: { contents: `export { collectVbaColors } from './src/vbaColors';
    export { resolveEnumValues } from './src/analyzer/completion/assignmentValueCompletion';`,
    resolveDir: process.cwd(), loader: 'ts' },
  bundle: true, write: false, platform: 'node', format: 'esm', logLevel: 'silent',
});
const { collectVbaColors, resolveEnumValues } = await import(`data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].text).toString('base64')}`);
resolveEnumValues('', 'XlHAlign', {}); // exclude first host-model initialization
for (const count of [200, 800, 2000]) {
  const source = `Sub T()\n${'Me.BackColor = 255\n'.repeat(count)}End Sub`;
  const start = performance.now();
  const colors = collectVbaColors(source);
  const cold = performance.now() - start;
  const warmSamples = [];
  for (let i = 0; i < 7; i++) {
    const warmStart = performance.now();
    if (collectVbaColors(source).length !== count) throw new Error('Missing color decorations');
    warmSamples.push(performance.now() - warmStart);
  }
  warmSamples.sort((a, b) => a - b);
  const enumStart = performance.now();
  for (let i = 0; i < 30; i++) resolveEnumValues(source, 'XlHAlign', {});
  console.log(JSON.stringify({ statements: count, colors: colors.length, coldMs: +cold.toFixed(2), warmMedianMs: +warmSamples[3].toFixed(2), warmMinMs: +warmSamples[0].toFixed(2), enumLookup30Ms: +(performance.now() - enumStart).toFixed(2) }));
}
