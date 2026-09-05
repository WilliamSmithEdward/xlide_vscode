import { ACCESS_PAGE_SIZE, AccessFormatError, AccessPageType } from './accessFormat';
import { AccessPageStore } from './accessPageStore';

/**
 * B-tree index pages, and taking one entry out of a tree.
 *
 * A page (type 3 node, type 4 leaf) holds a bit mask over its entry area
 * whose set bits mark the END of each entry, the first starting at 0. Every
 * entry after the first is stored without the leading bytes it shares with the
 * first, that many being the prefix length at 0x18. An entry is the encoded
 * key, then the row's home slot as a big-endian three-byte page and one-byte
 * row, and on a node page the big-endian child page.
 *
 * Entries are ordered by their full bytes, key then row pointer. A page is
 * rewritten compactly: entries after the first drop the bytes they share with
 * it. A leaf that fills while an entry is appended at its end stays full and
 * the next leaf starts with the new entry; an insert in the middle splits the
 * page in half. The root page number never changes: a splitting root turns
 * into a node whose two children are fresh pages.
 *
 * Ported from pyOpenVBA's `_index.py` and `_btree.py`.
 */

const OFFSET_FREE_SPACE = 0x02;
const OFFSET_OWNER = 0x04;
const OFFSET_PREV = 0x0c;
const OFFSET_NEXT = 0x10;
const OFFSET_TAIL = 0x14;
const OFFSET_PREFIX_LENGTH = 0x18;
/** The page's height above the leaves. */
const OFFSET_LEVEL = 0x1a;
const OFFSET_ENTRY_MASK = 0x1b;
const SIZE_ENTRY_MASK = 453;
const OFFSET_ENTRIES = OFFSET_ENTRY_MASK + SIZE_ENTRY_MASK;
const ENTRY_AREA = ACCESS_PAGE_SIZE - OFFSET_ENTRIES;

export interface AccessIndexEntry {
	key: Buffer;
	/** The home page of the row the entry points at. */
	page: number;
	row: number;
	/** The child page, on a node entry. */
	child?: number;
}

export interface AccessIndexPage {
	number: number;
	isLeaf: boolean;
	owner: number;
	prev: number;
	next: number;
	tail: number;
	prefixLength: number;
	entries: AccessIndexEntry[];
}

export function parseIndexPage(store: AccessPageStore, number: number): AccessIndexPage {
	const raw = store.read(number);
	const kind = raw[0];
	if (kind !== AccessPageType.IndexNode && kind !== AccessPageType.IndexLeaf) {
		throw new AccessFormatError(`Page ${number} is type 0x${kind.toString(16)}, not an index page.`);
	}
	const isLeaf = kind === AccessPageType.IndexLeaf;
	const prefixLength = raw.readUInt16LE(OFFSET_PREFIX_LENGTH);
	const entries: AccessIndexEntry[] = [];
	let start = 0;
	let first: Buffer | undefined;
	for (let byteIndex = 0; byteIndex < SIZE_ENTRY_MASK; byteIndex += 1) {
		const byte = raw[OFFSET_ENTRY_MASK + byteIndex];
		if (byte === 0) {
			continue;
		}
		for (let bit = 0; bit < 8; bit += 1) {
			if ((byte & (1 << bit)) === 0) {
				continue;
			}
			const end = byteIndex * 8 + bit;
			if (end <= start || end > ENTRY_AREA) {
				throw new AccessFormatError(
					`Index page ${number}: an entry boundary at ${end} after ${start} is impossible.`,
				);
			}
			const stored = raw.subarray(OFFSET_ENTRIES + start, OFFSET_ENTRIES + end);
			let full: Buffer;
			if (first === undefined) {
				first = stored;
				full = stored;
			} else {
				if (prefixLength > first.length) {
					throw new AccessFormatError(
						`Index page ${number}: the prefix of ${prefixLength} is longer than its first entry.`,
					);
				}
				full = Buffer.concat([first.subarray(0, prefixLength), stored]);
			}
			entries.push(splitEntry(full, isLeaf, number));
			start = end;
		}
	}
	return {
		number,
		isLeaf,
		owner: raw.readUInt32LE(OFFSET_OWNER),
		prev: raw.readUInt32LE(OFFSET_PREV),
		next: raw.readUInt32LE(OFFSET_NEXT),
		tail: raw.readUInt32LE(OFFSET_TAIL),
		prefixLength,
		entries,
	};
}

function splitEntry(full: Buffer, isLeaf: boolean, number: number): AccessIndexEntry {
	const trailer = isLeaf ? 4 : 8;
	if (full.length < trailer) {
		throw new AccessFormatError(
			`Index page ${number}: an entry of ${full.length} bytes has no row pointer.`,
		);
	}
	const at = full.length - trailer;
	return {
		key: Buffer.from(full.subarray(0, at)),
		page: full.readUIntBE(at, 3),
		row: full[at + 3],
		...(isLeaf ? {} : { child: full.readUInt32BE(full.length - 4) }),
	};
}

/** The leftmost leaf, reached by descending through first children. */
function firstLeaf(store: AccessPageStore, root: number): AccessIndexPage {
	const seen = new Set<number>();
	let page = parseIndexPage(store, root);
	while (!page.isLeaf) {
		if (seen.has(page.number)) {
			throw new AccessFormatError(`The index rooted at ${root} loops through page ${page.number}.`);
		}
		seen.add(page.number);
		const child = page.entries.length > 0 ? page.entries[0].child : page.tail;
		if (!child) {
			throw new AccessFormatError(`Index node ${page.number} has no children.`);
		}
		page = parseIndexPage(store, child);
	}
	return page;
}

/** Every leaf page in key order, following the next pointers. */
export function leafPages(store: AccessPageStore, root: number): AccessIndexPage[] {
	const out: AccessIndexPage[] = [];
	const seen = new Set<number>();
	let page = firstLeaf(store, root);
	for (;;) {
		if (seen.has(page.number)) {
			throw new AccessFormatError(`The index rooted at ${root} loops through page ${page.number}.`);
		}
		seen.add(page.number);
		out.push(page);
		if (!page.next) {
			return out;
		}
		page = parseIndexPage(store, page.next);
	}
}

/** An entry as stored before prefix removal. */
function entryBytes(entry: AccessIndexEntry): Buffer {
	const pointer = Buffer.alloc(4);
	pointer.writeUIntBE(entry.page, 0, 3);
	pointer[3] = entry.row;
	if (entry.child === undefined) {
		return Buffer.concat([entry.key, pointer]);
	}
	const child = Buffer.alloc(4);
	child.writeUInt32BE(entry.child, 0);
	return Buffer.concat([entry.key, pointer, child]);
}

/** What entries are ordered by: the key, then the row pointer. */
function sortKey(entry: AccessIndexEntry): Buffer {
	const pointer = Buffer.alloc(4);
	pointer.writeUIntBE(entry.page, 0, 3);
	pointer[3] = entry.row;
	return Buffer.concat([entry.key, pointer]);
}

/**
 * Bytes the entry area needs. `prefix` is the prefix length in force, shrunk
 * when the entries no longer share it; undefined means the full common prefix.
 */
function storedSize(entries: AccessIndexEntry[], prefix?: number): number {
	const raw = entries.map(entryBytes);
	const shared = commonPrefix(raw);
	const used = prefix === undefined ? shared : Math.min(prefix, shared);
	return raw.reduce((sum, item) => sum + item.length, 0) - used * Math.max(raw.length - 1, 0);
}

function commonPrefix(items: Buffer[]): number {
	if (items.length < 2) {
		return 0;
	}
	const first = items[0];
	let length = 0;
	while (length < first.length
		&& items.every((item) => item.length > length && item[length] === first[length])) {
		length += 1;
	}
	return length;
}

interface SerializeOptions {
	isLeaf: boolean;
	owner: number;
	prev: number;
	next: number;
	tail: number;
	level: number;
	/** The prefix length to keep; undefined computes the full common prefix. */
	prefix?: number;
	/** The page's previous content, whose bytes past the entries stay put. */
	base?: Buffer;
}

function serializeIndexPage(entries: AccessIndexEntry[], options: SerializeOptions): Buffer {
	const raw = options.base ? Buffer.from(options.base) : Buffer.alloc(ACCESS_PAGE_SIZE);
	raw[0] = options.isLeaf ? AccessPageType.IndexLeaf : AccessPageType.IndexNode;
	raw[1] = 0x01;
	raw.writeUInt32LE(options.owner, OFFSET_OWNER);
	raw.writeUInt32LE(options.prev, OFFSET_PREV);
	raw.writeUInt32LE(options.next, OFFSET_NEXT);
	raw.writeUInt32LE(options.tail, OFFSET_TAIL);
	raw[OFFSET_LEVEL] = options.level;
	raw.fill(0, OFFSET_ENTRY_MASK, OFFSET_ENTRIES);
	const stored = entries.map(entryBytes);
	const shared = commonPrefix(stored);
	const prefix = options.prefix === undefined ? shared : Math.min(options.prefix, shared);
	raw.writeUInt16LE(prefix, OFFSET_PREFIX_LENGTH);
	let position = 0;
	stored.forEach((item, index) => {
		const body = index === 0 ? item : item.subarray(prefix);
		const end = position + body.length;
		if (end > ENTRY_AREA) {
			throw new AccessFormatError(`${entries.length} index entries do not fit one page.`);
		}
		body.copy(raw, OFFSET_ENTRIES + position);
		raw[OFFSET_ENTRY_MASK + (end >> 3)] |= 1 << (end % 8);
		position = end;
	});
	raw.writeUInt16LE(ENTRY_AREA - position, OFFSET_FREE_SPACE);
	return raw;
}

/** One page on the path from the root, with the position taken in it. */
interface Step {
	page: AccessIndexPage;
	position: number;
}

/**
 * One index's tree, open for deletion. The root page number never changes, so
 * the table definition needs no patch when the tree is edited.
 */
export class AccessBTree {
	constructor(
		private readonly store: AccessPageStore,
		private readonly root: number,
		private readonly owner: number,
		/** Hands out a fresh page, already registered with the index's map. */
		private readonly allocate: () => number = () => {
			throw new AccessFormatError('This index tree was opened without a page allocator.');
		},
	) {}

	/** The leaf entry pointing at a row, or undefined when the row is absent. */
	entryOf(page: number, row: number): AccessIndexEntry | undefined {
		for (const leaf of leafPages(this.store, this.root)) {
			const found = leaf.entries.find((entry) => entry.page === page && entry.row === row);
			if (found) {
				return found;
			}
		}
		return undefined;
	}

	/**
	 * Remove the entry for a row. The engine leaves an emptied leaf in place
	 * and merges nothing, so this only rewrites the leaf and fixes the
	 * separators above it.
	 */
	delete(entry: AccessIndexEntry): void {
		const probe = sortKey(entry);
		const path = this.descend(probe);
		const leaf = path[path.length - 1];
		const entries = [...leaf.page.entries];
		if (leaf.position >= entries.length
			|| !sortKey(entries[leaf.position]).equals(probe)) {
			throw new AccessFormatError(
				`The index has no entry for row (${entry.page}, ${entry.row}) under this key.`,
			);
		}
		entries.splice(leaf.position, 1);
		this.rewrite(leaf.page, entries);
		if (entries.length > 0) {
			this.refreshSeparators(path, entries[entries.length - 1]);
		}
	}

	/**
	 * Insert an entry, and say whether the tree held no other entry with this
	 * key - which is what the definition's distinct-key counter follows.
	 */
	insert(key: Buffer, rowPage: number, row: number): boolean {
		const entry: AccessIndexEntry = { key, page: rowPage, row };
		const probe = sortKey(entry);
		const path = this.descend(probe);
		const leaf = path[path.length - 1];
		const entries = [...leaf.page.entries];
		if (leaf.position < entries.length && sortKey(entries[leaf.position]).equals(probe)) {
			throw new AccessFormatError(
				`The index already holds row (${rowPage}, ${row}) under this key.`,
			);
		}
		const distinct = this.keyIsNew(leaf, entries, key);
		entries.splice(leaf.position, 0, entry);
		this.place(path, entries, leaf.position === entries.length - 1);
		return distinct;
	}

	/**
	 * Write the page at the end of `path` with `entries`: as it is while they
	 * fit under its current prefix, compressed with the full common prefix once
	 * they do not, and split when even that is not enough or when an entry does
	 * not carry the prefix the page already has.
	 */
	private place(path: Step[], entries: AccessIndexEntry[], appended: boolean): void {
		const step = path[path.length - 1];
		const prefix = step.page.prefixLength;
		if (prefix > 0 && commonPrefix(entries.map(entryBytes)) < prefix) {
			// A page's prefix is fixed once it has been compressed: an entry that
			// does not start with those bytes belongs on another page.
			this.split(path, entries, appended);
			return;
		}
		if (storedSize(entries, prefix) <= ENTRY_AREA) {
			this.rewrite(step.page, entries);
		} else if (storedSize(entries) <= ENTRY_AREA) {
			this.rewrite(step.page, entries, { recompress: true });
		} else {
			this.split(path, entries, appended);
			return;
		}
		if (step.page.isLeaf) {
			this.refreshSeparators(path, entries[entries.length - 1]);
		}
	}

	/**
	 * Equal keys are adjacent, so only the neighbours of the insertion point can
	 * match - including across the leaf boundary.
	 */
	private keyIsNew(leaf: Step, entries: AccessIndexEntry[], key: Buffer): boolean {
		if (leaf.position > 0) {
			if (entries[leaf.position - 1].key.equals(key)) {
				return false;
			}
		} else if (leaf.page.prev) {
			const previous = this.load(leaf.page.prev);
			const last = previous.entries[previous.entries.length - 1];
			if (last && last.key.equals(key)) {
				return false;
			}
		}
		if (leaf.position < entries.length) {
			if (entries[leaf.position].key.equals(key)) {
				return false;
			}
		} else if (leaf.page.next) {
			const following = this.load(leaf.page.next);
			if (following.entries.length > 0 && following.entries[0].key.equals(key)) {
				return false;
			}
		}
		return true;
	}

	private divide(
		entries: AccessIndexEntry[],
		appended: boolean,
	): [AccessIndexEntry[], AccessIndexEntry[]] {
		if (appended) {
			return [entries.slice(0, -1), entries.slice(-1)];
		}
		const sizes = entries.map((entry) => entryBytes(entry).length);
		const total = sizes.reduce((sum, size) => sum + size, 0);
		let running = 0;
		let cut = Math.floor(entries.length / 2);
		for (let i = 0; i < sizes.length; i += 1) {
			running += sizes[i];
			if (running * 2 >= total) {
				cut = i + 1;
				break;
			}
		}
		cut = Math.max(1, Math.min(cut, entries.length - 1));
		return [entries.slice(0, cut), entries.slice(cut)];
	}

	/** Split the page at the end of `path` and register the halves above it. */
	private split(path: Step[], entries: AccessIndexEntry[], appended: boolean): void {
		const step = path[path.length - 1];
		const page = step.page;
		const [left, right] = this.divide(entries, appended);
		if (page.number === this.root) {
			this.splitRoot(page, left, right);
			return;
		}
		const level = this.store.read(page.number)[OFFSET_LEVEL];
		const rightNumber = this.allocate();
		let leftLast: AccessIndexEntry;
		let rightLast: AccessIndexEntry;
		if (page.isLeaf) {
			this.writePage(rightNumber, right, {
				isLeaf: true, prev: page.number, next: page.next, tail: 0, level: 0,
			});
			if (page.next) {
				const following = this.load(page.next);
				this.rewrite(following, following.entries, { prev: rightNumber });
			}
			this.rewrite(page, left, { next: rightNumber });
			leftLast = left[left.length - 1];
			rightLast = right[right.length - 1];
		} else {
			// The last left entry becomes the left node's tail child; the right
			// node keeps the original tail.
			const leftTail = left[left.length - 1].child ?? 0;
			this.writePage(rightNumber, right, {
				isLeaf: false, prev: 0, next: 0, tail: page.tail, level,
			});
			this.writePage(page.number, left.slice(0, -1), {
				isLeaf: false, prev: 0, next: 0, tail: leftTail, level,
			});
			const last = left[left.length - 1];
			leftLast = { key: last.key, page: last.page, row: last.row };
			rightLast = this.subtreeLast(rightNumber);
		}
		this.registerSplit(path.slice(0, -1), page.number, leftLast, rightNumber, rightLast);
	}

	private splitRoot(
		root: AccessIndexPage,
		left: AccessIndexEntry[],
		right: AccessIndexEntry[],
	): void {
		const level = this.store.read(root.number)[OFFSET_LEVEL];
		const leftNumber = this.allocate();
		const rightNumber = this.allocate();
		// The left half takes over the root's page image, so whatever the root
		// held beyond its entries travels with it and the new node keeps it.
		const image = Buffer.from(this.store.read(root.number));
		if (root.isLeaf) {
			this.writePage(leftNumber, left, {
				isLeaf: true, prev: 0, next: rightNumber, tail: 0, level: 0, base: image,
			});
			this.writePage(rightNumber, right, {
				isLeaf: true, prev: leftNumber, next: 0, tail: 0, level: 0,
			});
		} else {
			const leftTail = left[left.length - 1].child ?? 0;
			this.writePage(leftNumber, left.slice(0, -1), {
				isLeaf: false, prev: 0, next: 0, tail: leftTail, level, base: image,
			});
			this.writePage(rightNumber, right, {
				isLeaf: false, prev: 0, next: 0, tail: root.tail, level,
			});
		}
		const last = left[left.length - 1];
		const separator: AccessIndexEntry = {
			key: last.key, page: last.page, row: last.row, child: leftNumber,
		};
		// The new root is one entry with no prefix, over what the page held.
		this.store.write(root.number, serializeIndexPage([separator], {
			isLeaf: false, owner: this.owner, prev: 0, next: 0, tail: rightNumber,
			level: level + 1, prefix: 0, base: image,
		}));
	}

	/**
	 * `leftNumber`, an existing child of the node at the end of `path`, split off
	 * `rightNumber`; give the node an entry for each half, splitting the node too
	 * when it overflows.
	 */
	private registerSplit(
		path: Step[],
		leftNumber: number,
		leftLast: AccessIndexEntry,
		rightNumber: number,
		rightLast: AccessIndexEntry,
	): void {
		const step = path[path.length - 1];
		const node = step.page;
		const entries = [...node.entries];
		const separator: AccessIndexEntry = {
			key: leftLast.key, page: leftLast.page, row: leftLast.row, child: leftNumber,
		};
		let tail = node.tail;
		let appended: boolean;
		if (step.position < entries.length) {
			entries[step.position] = {
				key: rightLast.key, page: rightLast.page, row: rightLast.row, child: rightNumber,
			};
			entries.splice(step.position, 0, separator);
			appended = false;
		} else {
			entries.push(separator);
			tail = rightNumber;
			appended = true;
		}
		if (storedSize(entries, node.prefixLength) <= ENTRY_AREA) {
			this.rewrite(node, entries, { tail });
			return;
		}
		if (storedSize(entries) <= ENTRY_AREA) {
			this.rewrite(node, entries, { tail, recompress: true });
			return;
		}
		const grown: AccessIndexPage = { ...node, tail, entries };
		this.split(
			[...path.slice(0, -1), { page: grown, position: step.position }], entries, appended,
		);
	}

	/** The greatest entry under a subtree, which is its separator above. */
	private subtreeLast(number: number): AccessIndexEntry {
		let page = this.load(number);
		while (!page.isLeaf) {
			const child = page.tail ? page.tail : (page.entries[page.entries.length - 1].child ?? 0);
			page = this.load(child);
		}
		if (page.entries.length === 0) {
			throw new AccessFormatError(`Index leaf ${page.number} is empty.`);
		}
		return page.entries[page.entries.length - 1];
	}

	private writePage(
		number: number,
		entries: AccessIndexEntry[],
		options: {
			isLeaf: boolean; prev: number; next: number; tail: number; level: number; base?: Buffer;
		},
	): void {
		this.store.write(number, serializeIndexPage(entries, { ...options, owner: this.owner }));
	}

	private load(number: number): AccessIndexPage {
		const page = parseIndexPage(this.store, number);
		if (page.owner !== this.owner) {
			throw new AccessFormatError(
				`Index page ${number} belongs to table definition ${page.owner}, not ${this.owner}.`,
			);
		}
		return page;
	}

	private descend(probe: Buffer): Step[] {
		const path: Step[] = [];
		const seen = new Set<number>();
		let number = this.root;
		for (;;) {
			if (seen.has(number)) {
				throw new AccessFormatError(
					`The index rooted at ${this.root} loops through page ${number}.`,
				);
			}
			seen.add(number);
			const page = this.load(number);
			let position = 0;
			while (position < page.entries.length
				&& sortKey(page.entries[position]).compare(probe) < 0) {
				position += 1;
			}
			path.push({ page, position });
			if (page.isLeaf) {
				return path;
			}
			const child = position < page.entries.length ? page.entries[position].child : page.tail;
			if (!child) {
				throw new AccessFormatError(`Index node ${number} has no child to descend into.`);
			}
			number = child;
		}
	}

	/**
	 * Rewrite an existing page over its old bytes, keeping its prefix length
	 * (shrunk when the entries no longer share it).
	 */
	private rewrite(
		page: AccessIndexPage,
		entries: AccessIndexEntry[],
		overrides: { tail?: number; prev?: number; next?: number; recompress?: boolean } = {},
	): void {
		this.store.write(page.number, serializeIndexPage(entries, {
			isLeaf: page.isLeaf,
			owner: this.owner,
			prev: overrides.prev === undefined ? page.prev : overrides.prev,
			next: overrides.next === undefined ? page.next : overrides.next,
			tail: overrides.tail === undefined ? page.tail : overrides.tail,
			level: this.store.read(page.number)[OFFSET_LEVEL],
			...(overrides.recompress ? {} : { prefix: page.prefixLength }),
			base: this.store.read(page.number),
		}));
	}

	/** Keep every ancestor's separator equal to the greatest entry under it. */
	private refreshSeparators(path: Step[], childLast: AccessIndexEntry): void {
		for (let depth = path.length - 2; depth >= 0; depth -= 1) {
			const step = path[depth];
			if (step.position >= step.page.entries.length) {
				return; // the tail child has no separator
			}
			const current = step.page.entries[step.position];
			if (current.key.equals(childLast.key)
				&& current.page === childLast.page && current.row === childLast.row) {
				return;
			}
			const entries = [...step.page.entries];
			entries[step.position] = {
				key: childLast.key,
				page: childLast.page,
				row: childLast.row,
				...(current.child === undefined ? {} : { child: current.child }),
			};
			this.rewrite(step.page, entries);
			if (step.position !== entries.length - 1) {
				return;
			}
		}
	}
}
