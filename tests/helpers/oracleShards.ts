// The comment-lines sweep analyzes every oracle case twice, and the
// wrapped-in-If sweep every case it can wrap. As single tests they took
// 84 s and 64 s, each file run alone here on 2026-10-16. One test runs on
// one worker, so each sweep is split across shard files that vitest runs
// side by side: shard i takes the cases whose index modulo the count is i,
// which spreads the slow cases and covers each case once.
// oracleShards.test.ts checks that every shard file is there and runs its
// own slice.
export const oracleShardCount = 4;

export function oracleShardCases<T>(cases: readonly T[], shard: number): T[] {
	return cases.filter((_, index) => index % oracleShardCount === shard);
}
