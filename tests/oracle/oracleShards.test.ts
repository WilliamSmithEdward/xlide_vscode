import { readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import { expect, it } from 'vitest';
import { oracleShardCount } from '../helpers/oracleShards';

it('has every oracle sweep shard file, each running its own slice', () => {
	const tests = join(process.cwd(), 'tests', 'oracle');
	for (const sweep of ['oracleCommentLines', 'oracleWrappedInIf']) {
		expect(readdirSync(tests).filter((file) => file.startsWith(`${sweep}.shard`)).length).toBe(oracleShardCount);
		// A missing file throws; one that ran another's slice would leave a
		// quarter of the oracle unchecked.
		for (let index = 0; index < oracleShardCount; index++) {
			expect(readFileSync(join(tests, `${sweep}.shard${index + 1}.test.ts`), 'utf8')).toContain(`Shard(${index});`);
		}
	}
});
