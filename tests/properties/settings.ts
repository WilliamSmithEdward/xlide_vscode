import type fc from 'fast-check';

// How hard the property tests in this folder push. `npm test` runs each
// property a hundred times. The Fuzz workflow sets XLIDE_PROPERTY_RUNS to run
// them twenty thousand times when a reader changes and two hundred thousand
// times daily. A failure prints the smallest input that breaks the property
// and its seed; XLIDE_PROPERTY_SEED replays it.

export const PROPERTY_RUNS = Number(process.env.XLIDE_PROPERTY_RUNS ?? 100);

const SEED = process.env.XLIDE_PROPERTY_SEED === undefined ? undefined : Number(process.env.XLIDE_PROPERTY_SEED);

export function propertySettings<T>(numRuns = PROPERTY_RUNS): fc.Parameters<T> {
	return {
		numRuns,
		includeErrorInReport: true,
		...(SEED === undefined ? {} : { seed: SEED }),
	};
}
