// A form's whole designer storage as one recursive model.
//
// The storage tree mirrors the control tree: the form's storage holds `f`,
// `o`, and one `iNN` child storage per container control, where NN is the
// container site's ID ([MS-OFORMS] 2.1.2.2.2) - a Frame holds its own f/o,
// a MultiPage holds f (whose sites are its Pages), o (its TabStrip record),
// `x` (page bookkeeping, preserved verbatim), and one `iNN` per Page.

import { Cfb } from '../cfb';
import {
	parseFormStream,
	serializeFormStream,
	parseObjectStream,
	serializeObjectStream,
	siteId,
	siteIsContainer,
	siteCacheIndex,
	type FormStreamModel,
	type ObjectStreamEntry,
	type SiteModel,
} from './formStream';
import { LATIN1_CODEC, type OformsTextCodec } from './records';

const COMPOBJ_STREAM = '\x01CompObj';

export interface FormPackage {
	form: FormStreamModel;
	entries: ObjectStreamEntry[];
	/** Container site ID -> the nested package behind its `iNN` storage. */
	containers: Map<number, FormPackage>;
	/** MultiPage page bookkeeping (`x`), preserved verbatim. */
	xRaw?: Buffer;
	/** The storage's CompObj, kept so a copy can seed a new sibling. */
	compObjRaw?: Buffer;
}

/** The `iNN` storage name for a container site ID ([MS-OFORMS] 2.1.2.2.2). */
export function containerStorageName(id: number): string {
	return `i${id < 10 ? `0${id}` : String(id)}`;
}

export function parseFormPackage(
	cfb: Cfb,
	path: readonly string[],
	codec: OformsTextCodec = LATIN1_CODEC,
): FormPackage {
	const f = cfb.getStreamAtPath(path, 'f');
	const o = cfb.getStreamAtPath(path, 'o');
	const form = parseFormStream(f, codec);
	const entries = parseObjectStream(o, form.sites, codec);
	const pkg: FormPackage = { form, entries, containers: new Map() };
	if (cfb.hasStreamAtPath(path, 'x')) {
		pkg.xRaw = cfb.getStreamAtPath(path, 'x');
	}
	if (cfb.hasStreamAtPath(path, COMPOBJ_STREAM)) {
		pkg.compObjRaw = cfb.getStreamAtPath(path, COMPOBJ_STREAM);
	}
	for (const site of form.sites) {
		if (!siteIsContainer(site)) { continue; }
		const child = containerStorageName(siteId(site));
		if (!cfb.hasStoragePath([...path, child])) {
			throw new RangeError(`container ${child} has no storage under ${path.join('/')}`);
		}
		pkg.containers.set(siteId(site), parseFormPackage(cfb, [...path, child], codec));
	}
	return pkg;
}

/**
 * Writes the package back: `f` and `o` always, `x` when carried, child
 * storages recursively - creating the storage (with a CompObj copied from the
 * parent) for a container the model gained, and removing the storage for one
 * it lost.
 */
export function writeFormPackage(
	cfb: Cfb,
	path: readonly string[],
	pkg: FormPackage,
	codec: OformsTextCodec = LATIN1_CODEC,
): void {
	// Children first: a container's own f/o must land before the parent's site
	// table (ObjectStreamSize refreshes) is serialized. Order does not affect
	// bytes here, but removing stale storages before adding avoids collisions.
	const wanted = new Set<string>();
	for (const [id, child] of pkg.containers) {
		wanted.add(containerStorageName(id).toLowerCase());
		const childPath = [...path, containerStorageName(id)];
		if (!cfb.hasStoragePath(childPath)) {
			cfb.addStorageAtPath(path, containerStorageName(id));
			const compObj = child.compObjRaw ?? pkg.compObjRaw;
			if (compObj) {
				cfb.setStreamAtPath(childPath, COMPOBJ_STREAM, compObj);
			}
		}
		writeFormPackage(cfb, childPath, child, codec);
	}
	for (const stale of cfb.listStoragesAtPath(path)) {
		if (/^i\d\d+$/i.test(stale) && !wanted.has(stale.toLowerCase())) {
			cfb.removeStorageAtPath([...path, stale]);
		}
	}

	// The o stream refreshes every non-container site's ObjectStreamSize, so
	// it must serialize before f.
	const o = serializeObjectStream(pkg.entries, codec);
	const f = serializeFormStream(pkg.form, codec);
	cfb.setStreamAtPath(path, 'o', o);
	cfb.setStreamAtPath(path, 'f', f);
	if (pkg.xRaw) {
		cfb.setStreamAtPath(path, 'x', pkg.xRaw);
	}
}

/** Depth-first walk over every package, parents before children. */
export function walkPackages(
	pkg: FormPackage,
	visit: (pkg: FormPackage, parentSite: SiteModel | undefined) => void,
	parentSite?: SiteModel,
): void {
	visit(pkg, parentSite);
	for (const site of pkg.form.sites) {
		if (!siteIsContainer(site)) { continue; }
		const child = pkg.containers.get(siteId(site));
		if (child) { walkPackages(child, visit, site); }
	}
}

/** The kind name a site's cache index answers to, for markup and display. */
export function controlKindOfSite(site: SiteModel, record?: { values: Map<string, number> }): string {
	switch (siteCacheIndex(site)) {
		case 7: return 'Page';
		case 12: return 'Image';
		case 14: return 'Frame';
		case 16: return 'SpinButton';
		case 17: return 'CommandButton';
		case 18: return 'TabStrip';
		case 21: return 'Label';
		case 23: return 'TextBox';
		case 24: return 'ListBox';
		case 25: return 'ComboBox';
		case 26: return 'CheckBox';
		case 27: return 'OptionButton';
		case 28: return 'ToggleButton';
		case 47: return 'ScrollBar';
		case 57: return 'MultiPage';
		case 15: {
			// Generic MorphData: DisplayStyle decides, defaulting to Text.
			const style = record?.values.get('DisplayStyle') ?? 1;
			return ({ 1: 'TextBox', 2: 'ListBox', 3: 'ComboBox', 4: 'CheckBox', 5: 'OptionButton', 6: 'ToggleButton', 7: 'ComboBox' } as Record<number, string>)[style] ?? 'TextBox';
		}
		default: return 'ActiveX';
	}
}
