// Canvas gestures, applied to the designer model: move, resize, add, remove.
//
// These are the same mutations the markup diff performs, addressed by control
// name instead of by document - VBA requires control names to be unique
// form-wide, so a name finds its control wherever it nests. Additions run
// through the same authoring path the markup uses, which is the path live
// Excel verified: per-kind recipes, tree-global IDs, the cookie rules.

import { pointsToHimetric } from './bytes';
import { siteName, siteId, siteIsContainer, siteCacheIndex, type SiteModel } from './formStream';
import { walkPackages, type FormPackage } from './formPackage';
import {
	addControlForDesigner,
	FormMarkupError,
	type MarkupElement,
} from './markup';

export interface ControlLocation {
	pkg: FormPackage;
	site: SiteModel;
	entry: FormPackage['entries'][number];
}

/** Finds a control by name anywhere in the form tree. */
export function findControl(root: FormPackage, name: string): ControlLocation | undefined {
	let found: ControlLocation | undefined;
	walkPackages(root, (pkg) => {
		if (found) { return; }
		for (const entry of pkg.entries) {
			if (siteName(entry.site).toLowerCase() === name.toLowerCase()) {
				found = { pkg, site: entry.site, entry };
				return;
			}
		}
	});
	return found;
}

/** Finds a container surface (a Frame or Page package) by its name; '' is the root. */
export function findSurface(root: FormPackage, containerName: string): FormPackage | undefined {
	if (!containerName) { return root; }
	let found: FormPackage | undefined;
	walkPackages(root, (pkg) => {
		if (found) { return; }
		for (const site of pkg.form.sites) {
			if (!siteIsContainer(site)) { continue; }
			if (siteName(site).toLowerCase() === containerName.toLowerCase()) {
				found = pkg.containers.get(siteId(site));
				return;
			}
		}
	});
	return found;
}

export interface GeometryPt {
	left?: number;
	top?: number;
	width?: number;
	height?: number;
}

/**
 * Moves and/or resizes one control, in points. Position lands on the site;
 * size lands on the record, or on the container's own form for a Frame,
 * MultiPage, or Page.
 */
export function setControlGeometry(root: FormPackage, name: string, geometry: GeometryPt): string[] {
	const location = findControl(root, name);
	if (!location) {
		throw new FormMarkupError(0, `no control named ${name}`);
	}
	const applied: string[] = [];
	const { site, entry, pkg } = location;
	if (geometry.left !== undefined || geometry.top !== undefined) {
		const current = site.position ?? { left: 0, top: 0 };
		const next = {
			left: geometry.left !== undefined ? pointsToHimetric(geometry.left) : current.left,
			top: geometry.top !== undefined ? pointsToHimetric(geometry.top) : current.top,
		};
		if (next.left !== current.left || next.top !== current.top) {
			site.position = next;
			site.mask = (site.mask | (1 << 8)) >>> 0;
			applied.push(`position of ${siteName(site)}`);
		}
	}
	if (geometry.width !== undefined || geometry.height !== undefined) {
		const target = entry.kind === 'record'
			? { record: entry.record, key: 'Size' as const }
			: undefined;
		if (target) {
			const size = target.record.sizes.get(target.key) ?? { width: 0, height: 0 };
			const next = {
				width: geometry.width !== undefined ? pointsToHimetric(geometry.width) : size.width,
				height: geometry.height !== undefined ? pointsToHimetric(geometry.height) : size.height,
			};
			if (next.width !== size.width || next.height !== size.height) {
				target.record.sizes.set(target.key, next);
				applied.push(`size of ${siteName(site)}`);
			}
		} else {
			const inner = pkg.containers.get(siteId(site));
			const record = inner?.form.record;
			const size = record?.sizes.get('DisplayedSize');
			if (record && size) {
				const next = {
					width: geometry.width !== undefined ? pointsToHimetric(geometry.width) : size.width,
					height: geometry.height !== undefined ? pointsToHimetric(geometry.height) : size.height,
				};
				if (next.width !== size.width || next.height !== size.height) {
					record.sizes.set('DisplayedSize', next);
					const logical = record.sizes.get('LogicalSize');
					if (logical && (logical.width !== 0 || logical.height !== 0)) {
						record.sizes.set('LogicalSize', next);
					}
					applied.push(`size of ${siteName(site)}`);
				}
			}
		}
	}
	return applied;
}

/** A fresh name for a kind, unique across the whole tree: Label1, Label2... */
export function nextControlName(root: FormPackage, kind: string): string {
	const taken = new Set<string>();
	walkPackages(root, (pkg) => {
		for (const site of pkg.form.sites) { taken.add(siteName(site).toLowerCase()); }
	});
	for (let i = 1; ; i++) {
		const candidate = `${kind}${i}`;
		if (!taken.has(candidate.toLowerCase())) { return candidate; }
	}
}

/** The default size a new control of `kind` takes on the canvas, in points. */
export function defaultSizePt(kind: string): { width: number; height: number } {
	switch (kind) {
		case 'Label': return { width: 72, height: 12 };
		case 'CommandButton': case 'ToggleButton': return { width: 72, height: 24 };
		case 'CheckBox': case 'OptionButton': return { width: 90, height: 15 };
		case 'ComboBox': case 'TextBox': return { width: 96, height: 18 };
		case 'ListBox': return { width: 96, height: 60 };
		case 'Frame': return { width: 120, height: 90 };
		case 'Image': return { width: 72, height: 54 };
		case 'SpinButton': return { width: 13, height: 36 };
		case 'ScrollBar': return { width: 13, height: 90 };
		case 'TabStrip': return { width: 150, height: 90 };
		default: return { width: 72, height: 24 };
	}
}

/**
 * Adds a control of `kind` at (left, top) points on the named surface,
 * through the markup's own authoring path. Returns the new control's name.
 */
export function addControlAt(
	root: FormPackage,
	containerName: string,
	kind: string,
	leftPt: number,
	topPt: number,
): string {
	const surface = findSurface(root, containerName);
	if (!surface) {
		throw new FormMarkupError(0, `no container named ${containerName}`);
	}
	const name = nextControlName(root, kind);
	const size = defaultSizePt(kind);
	const attrs = new Map<string, string>([
		['Name', name],
		['Left', String(leftPt)],
		['Top', String(topPt)],
		['Width', String(size.width)],
		['Height', String(size.height)],
	]);
	if (['Label', 'CommandButton', 'ToggleButton', 'CheckBox', 'OptionButton', 'Frame'].includes(kind)) {
		attrs.set('Caption', name);
	}
	const element: MarkupElement = { tag: kind, attrs, children: [], line: 0 };
	addControlForDesigner(surface, element, root);
	return name;
}

/**
 * Removes one control by name. Pages are refused here - removing a page
 * moves four structures at once and belongs to the markup diff, where the
 * whole document states the intent.
 */
export function removeControl(root: FormPackage, name: string): void {
	const location = findControl(root, name);
	if (!location) {
		throw new FormMarkupError(0, `no control named ${name}`);
	}
	if (siteCacheIndex(location.site) === 7) {
		throw new FormMarkupError(0, `${name} is a Page; remove pages through the form markup`);
	}
	const { pkg, site } = location;
	pkg.form.sites = pkg.form.sites.filter((s) => s !== site);
	pkg.entries = pkg.entries.filter((e) => e.site !== site);
	pkg.containers.delete(siteId(site));
	pkg.form.sitesStructurallyChanged = true;
}

/** Resizes the form's own client area, in points. */
export function setFormSize(root: FormPackage, widthPt: number, heightPt: number): void {
	const record = root.form.record;
	const size = record.sizes.get('DisplayedSize');
	if (!size) { return; }
	const next = { width: pointsToHimetric(widthPt), height: pointsToHimetric(heightPt) };
	record.sizes.set('DisplayedSize', next);
	const logical = record.sizes.get('LogicalSize');
	if (logical && (logical.width !== 0 || logical.height !== 0)) {
		record.sizes.set('LogicalSize', next);
	}
}
