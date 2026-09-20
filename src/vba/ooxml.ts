// Reading and rewriting OOXML parts in place: the layer every host's shape
// reader and writer is built on.
//
// These packages are edited as text, not through a DOM. An Office file that
// a round trip reformats is one Office may re-save differently or refuse, so
// every edit here is a splice: the bytes outside the span being changed come
// back identical, attribute order is preserved, and a part nothing touched is
// never rewritten. That is also what makes the result diffable.
//
// The element walk is deliberately small - it counts nested elements of the
// same name and nothing else. It is not a parser: it assumes the
// well-formed, namespace-prefixed markup Office writes, where a prefix is
// bound once on the root and never rebound, which is true of every part
// Excel, Word and PowerPoint produce.

import { ZipArchive } from './zip';

export const CONTENT_TYPES = '[Content_Types].xml';

export function decodeXml(text: string): string {
	return text.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
		.replace(/&#x([0-9a-fA-F]+);/g, (_s, h: string) => String.fromCodePoint(parseInt(h, 16)))
		.replace(/&#(\d+);/g, (_s, d: string) => String.fromCodePoint(Number(d)))
		.replace(/&amp;/g, '&');
}

export function encodeXml(text: string): string {
	return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** An attribute of a start tag, decoded. */
export function attr(tag: string, name: string): string | undefined {
	const m = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*"([^"]*)"`).exec(tag);
	return m ? decodeXml(m[1]) : undefined;
}

/**
 * A start tag with an attribute set where it stands, added last when the tag
 * has none, or removed when `value` is undefined.
 */
export function withAttr(tag: string, name: string, value: string | undefined): string {
	const re = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*"[^"]*"`);
	if (value === undefined) { return tag.replace(re, ''); }
	const set = ` ${name}="${encodeXml(value)}"`;
	if (re.test(tag)) { return tag.replace(re, () => set); }
	const close = tag.endsWith('/>') ? tag.length - 2 : tag.length - 1;
	return tag.slice(0, close) + set + tag.slice(close);
}

export interface Span {
	start: number;
	/** Just past the start tag. */
	openEnd: number;
	end: number;
}

/** The element `name` starting at or after `from`, counting nested elements of the same name. */
export function findElement(xml: string, name: string, from = 0, until = xml.length): Span | undefined {
	const open = new RegExp(`<${name}(?=[\\s/>])`, 'g');
	open.lastIndex = from;
	const first = open.exec(xml);
	if (!first || first.index >= until) { return undefined; }
	const openEnd = xml.indexOf('>', first.index) + 1;
	if (xml[openEnd - 2] === '/') { return { start: first.index, openEnd, end: openEnd }; }
	const tag = new RegExp(`<${name}(?=[\\s/>])[^>]*?(/?)>|</${name}>`, 'g');
	tag.lastIndex = openEnd;
	let depth = 1;
	for (let m = tag.exec(xml); m; m = tag.exec(xml)) {
		if (m[0].startsWith('</')) {
			if (--depth === 0) { return { start: first.index, openEnd, end: m.index + m[0].length }; }
		} else if (m[1] !== '/') {
			depth++;
		}
	}
	return undefined;
}

/** The child elements of the element spanning [openEnd, close), in order. */
export function children(xml: string, openEnd: number, close: number): Array<Span & { name: string }> {
	const out: Array<Span & { name: string }> = [];
	let at = openEnd;
	for (;;) {
		const lt = xml.indexOf('<', at);
		if (lt < 0 || lt >= close || xml.startsWith('</', lt)) { return out; }
		if (xml.startsWith('<!--', lt) || xml.startsWith('<?', lt)) {
			at = xml.indexOf('>', lt) + 1;
			continue;
		}
		const name = /^<([\w:.-]+)/.exec(xml.slice(lt))![1];
		const span = findElement(xml, name, lt, close);
		if (!span) { return out; }
		out.push({ ...span, name });
		at = span.end;
	}
}

/** `xml` with the bytes of `span` replaced by `text`. */
export function splice(xml: string, span: { start: number; end: number }, text: string): string {
	return xml.slice(0, span.start) + text + xml.slice(span.end);
}

/**
 * Insert `text` as a child of the element spanning [openEnd, close), placed
 * so that the result still follows `order` - the schema's element sequence,
 * which Office enforces on load. Names not in `order` sort last.
 */
export function insertInOrder(
	xml: string,
	openEnd: number,
	close: number,
	order: readonly string[],
	name: string,
	text: string,
): string {
	const rank = (child: string): number => {
		const at = order.indexOf(child);
		return at < 0 ? order.length : at;
	};
	const mine = rank(name);
	const after = children(xml, openEnd, close).find((child) => rank(child.name) > mine);
	return splice(xml, { start: after?.start ?? close, end: after?.start ?? close }, text);
}

// ------------------------------------------------------------------ package

export interface Relationship {
	id: string;
	type: string;
	target: string;
	/** The target as a package path. */
	path: string;
}

export function relsPathOf(part: string): string {
	const slash = part.lastIndexOf('/');
	return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

/** A relationship target resolved against the folder of the part that holds it. */
export function resolveTarget(part: string, target: string): string {
	if (target.startsWith('/')) { return target.slice(1); }
	const segments = part.split('/').slice(0, -1);
	for (const segment of target.split('/')) {
		if (segment === '..') { segments.pop(); } else if (segment !== '.') { segments.push(segment); }
	}
	return segments.join('/');
}

/** A package path as a target relative to the folder of `part`. */
export function relativeTarget(part: string, path: string): string {
	const from = part.split('/').slice(0, -1);
	const to = path.split('/');
	let common = 0;
	while (common < from.length && common < to.length - 1 && from[common] === to[common]) { common++; }
	return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/');
}

/** An OOXML package: its parts, their relationships and the content types. */
export class Package {
	constructor(private readonly zip: ZipArchive) {}

	has(path: string): boolean {
		return this.zip.has(path);
	}

	read(path: string): string {
		return this.zip.read(path).toString('utf8');
	}

	write(path: string, text: string): void {
		this.zip.write(path, Buffer.from(text, 'utf8'));
	}

	relationships(part: string): Relationship[] {
		const rels = relsPathOf(part);
		if (!this.zip.has(rels)) { return []; }
		return [...this.read(rels).matchAll(/<Relationship\b[^>]*>/g)].map(([tag]) => ({
			id: attr(tag, 'Id') ?? '',
			type: attr(tag, 'Type') ?? '',
			target: attr(tag, 'Target') ?? '',
			path: resolveTarget(part, attr(tag, 'Target') ?? ''),
		}));
	}

	addRelationship(part: string, type: string, path: string): string {
		const rels = relsPathOf(part);
		const xml = this.zip.has(rels)
			? this.read(rels)
			: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
		const ids = [...xml.matchAll(/\bId="rId(\d+)"/g)].map((m) => Number(m[1]));
		const id = `rId${Math.max(0, ...ids) + 1}`;
		this.write(rels, xml.replace(/<\/Relationships>/,
			`<Relationship Id="${id}" Type="${type}" Target="${encodeXml(relativeTarget(part, path))}"/></Relationships>`));
		return id;
	}

	removeRelationship(part: string, id: string): void {
		const rels = relsPathOf(part);
		if (!this.zip.has(rels)) { return; }
		const xml = this.read(rels).replace(new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*/>`), '');
		if (/<Relationship\b/.test(xml)) {
			this.write(rels, xml);
		} else {
			this.zip.delete(rels);
		}
	}

	addOverride(path: string, contentType: string): void {
		const xml = this.read(CONTENT_TYPES);
		if (!xml.includes(`PartName="/${path}"`)) {
			this.write(CONTENT_TYPES, xml.replace(/<\/Types>/, `<Override PartName="/${path}" ContentType="${contentType}"/></Types>`));
		}
	}

	ensureDefault(extension: string, contentType: string): void {
		const xml = this.read(CONTENT_TYPES);
		if (!new RegExp(`<Default\\b[^>]*Extension="${extension}"`, 'i').test(xml)) {
			this.write(CONTENT_TYPES, xml.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${extension}" ContentType="${contentType}"/>`));
		}
	}

	/** Remove a part, its relationships, and every part only it referred to. */
	removePart(path: string): void {
		const targets = this.relationships(path).map((rel) => rel.path);
		this.zip.delete(path);
		this.zip.delete(relsPathOf(path));
		const types = this.read(CONTENT_TYPES).replace(new RegExp(`<Override\\b[^>]*PartName="/${path.replace(/[.[\]]/g, '\\$&')}"[^>]*/>`), '');
		this.write(CONTENT_TYPES, types);
		for (const target of targets) {
			if (this.zip.has(target) && !this.referenced(target)) {
				this.removePart(target);
			}
		}
	}

	/** A path free to take, as `xl/drawings/drawing3.xml` for base `xl/drawings/drawing` and `.xml`. */
	freePath(base: string, extension: string): string {
		for (let n = 1; ; n++) {
			if (!this.zip.has(`${base}${n}${extension}`)) { return `${base}${n}${extension}`; }
		}
	}

	private referenced(path: string): boolean {
		return this.zip.names().filter((name) => name.endsWith('.rels')).some((rels) => {
			const part = rels.replace(/_rels\/([^/]+)\.rels$/, '$1');
			return this.relationships(part).some((rel) => rel.path === path);
		});
	}

	names(): string[] {
		return this.zip.names();
	}
}

// ------------------------------------------------------------------ geometry

/**
 * English Metric Units per point. DrawingML measures in EMU; the Office
 * object model, and so every number a user has seen, is in points.
 */
export const EMU_PER_POINT = 12700;

export function emuToPoints(emu: number): number {
	// Two decimals: enough to round-trip any whole EMU a user typed as points,
	// without showing the noise of a drag.
	return Math.round((emu / EMU_PER_POINT) * 100) / 100;
}

export function pointsToEmu(points: number): number {
	return Math.round(points * EMU_PER_POINT);
}

/** A position and size in points, as every host's object model reports one. */
export interface PointBox {
	left?: number;
	top?: number;
	width?: number;
	height?: number;
}

/**
 * An a:xfrm read as points: a:off is the position and a:ext the size, and
 * either may be absent. Every DrawingML host spells this the same way, on a
 * slide and inside a Word canvas alike.
 */
export function xfrmBox(xml: string, xfrm: Span): PointBox {
	const inner = xml.slice(xfrm.start, xfrm.end);
	const off = /<a:off\b[^>]*>/.exec(inner)?.[0];
	const ext = /<a:ext\b[^>]*>/.exec(inner)?.[0];
	const box: PointBox = {};
	if (off) {
		box.left = emuToPoints(Number(attr(off, 'x') ?? 0));
		box.top = emuToPoints(Number(attr(off, 'y') ?? 0));
	}
	if (ext) {
		box.width = emuToPoints(Number(attr(ext, 'cx') ?? 0));
		box.height = emuToPoints(Number(attr(ext, 'cy') ?? 0));
	}
	return box;
}
