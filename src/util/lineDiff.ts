// A line diff and its unified rendering, for two versions of a module.
//
// Myers' O(ND) algorithm over lines, after the common head and tail are
// trimmed off: a module edit usually touches a few lines in the middle, so
// the search runs on the changed stretch, not the whole module. The unified
// text is what `git diff` prints, which is what an agent or a reader already
// knows how to read. No runtime dependency; the repository has none.

export type LineDiffOp =
	| { kind: 'equal'; line: string }
	| { kind: 'delete'; line: string }
	| { kind: 'insert'; line: string };

export interface UnifiedDiffOptions {
	/** The `---` header label, e.g. `Module1 (HEAD)`. */
	beforeLabel: string;
	/** The `+++` header label, e.g. `Module1 (current)`. */
	afterLabel: string;
	/** Unchanged lines shown around each change; git's default is 3. */
	context?: number;
}

/** Lines of a source, line endings dropped; an empty source has no lines. */
export function splitLines(source: string): string[] {
	if (source === '') {
		return [];
	}
	const lines = source.split(/\r\n|\r|\n/);
	// A trailing terminator ends the last line rather than starting an empty one.
	if (lines.length > 0 && lines[lines.length - 1] === '' && /[\r\n]$/.test(source)) {
		lines.pop();
	}
	return lines;
}

/** The edit script that turns `before` into `after`, one op per line. */
export function diffLines(before: readonly string[], after: readonly string[]): LineDiffOp[] {
	let head = 0;
	while (head < before.length && head < after.length && before[head] === after[head]) {
		head++;
	}
	let tail = 0;
	while (
		tail < before.length - head
		&& tail < after.length - head
		&& before[before.length - 1 - tail] === after[after.length - 1 - tail]
	) {
		tail++;
	}
	const ops: LineDiffOp[] = [];
	for (let i = 0; i < head; i++) {
		ops.push({ kind: 'equal', line: before[i] });
	}
	ops.push(...myers(before.slice(head, before.length - tail), after.slice(head, after.length - tail)));
	for (let i = before.length - tail; i < before.length; i++) {
		ops.push({ kind: 'equal', line: before[i] });
	}
	return ops;
}

/**
 * Myers' shortest edit script. The V arrays of every step are kept so the
 * script can be read back from the end; the changed stretch of a module is
 * short, so that memory is small.
 */
function myers(a: readonly string[], b: readonly string[]): LineDiffOp[] {
	const n = a.length;
	const m = b.length;
	if (n === 0) {
		return b.map((line) => ({ kind: 'insert', line }));
	}
	if (m === 0) {
		return a.map((line) => ({ kind: 'delete', line }));
	}
	const max = n + m;
	const offset = max + 1;
	// trace[d] is the furthest-reaching state before round d; round d reads
	// its neighbours from it and writes a fresh copy, so nothing is mutated
	// after it is recorded.
	const trace: Int32Array[] = [];
	let v = new Int32Array(2 * max + 3);
	let found = false;
	for (let d = 0; d <= max && !found; d++) {
		trace.push(v);
		const next = new Int32Array(v);
		for (let k = -d; k <= d; k += 2) {
			let x = k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1])
				? v[offset + k + 1]
				: v[offset + k - 1] + 1;
			let y = x - k;
			while (x < n && y < m && a[x] === b[y]) {
				x++;
				y++;
			}
			next[offset + k] = x;
			if (x >= n && y >= m) {
				found = true;
				break;
			}
		}
		v = next;
	}
	// Read the script back from the end: at each round, which neighbour the
	// path came from, then the diagonal it followed.
	const ops: LineDiffOp[] = [];
	let x = n;
	let y = m;
	for (let d = trace.length - 1; d >= 0; d--) {
		const before = trace[d];
		const k = x - y;
		const previousK = k === -d || (k !== d && before[offset + k - 1] < before[offset + k + 1])
			? k + 1
			: k - 1;
		const previousX = before[offset + previousK];
		const previousY = previousX - previousK;
		while (x > previousX && y > previousY) {
			x--;
			y--;
			ops.push({ kind: 'equal', line: a[x] });
		}
		if (d > 0) {
			if (x === previousX) {
				y--;
				ops.push({ kind: 'insert', line: b[y] });
			} else {
				x--;
				ops.push({ kind: 'delete', line: a[x] });
			}
		}
	}
	return ops.reverse();
}

/**
 * The diff of two sources as `git diff` would print it. An empty string when
 * the sources hold the same lines.
 */
export function unifiedDiff(before: string, after: string, options: UnifiedDiffOptions): string {
	const context = Math.max(0, options.context ?? 3);
	const ops = diffLines(splitLines(before), splitLines(after));
	if (!ops.some((op) => op.kind !== 'equal')) {
		return '';
	}
	const out: string[] = [`--- ${options.beforeLabel}`, `+++ ${options.afterLabel}`];
	let beforeLine = 1;
	let afterLine = 1;
	let index = 0;
	while (index < ops.length) {
		if (ops[index].kind === 'equal') {
			beforeLine++;
			afterLine++;
			index++;
			continue;
		}
		// A hunk: from `context` lines before the first change to `context`
		// lines after the last change of a run whose gaps are at most twice
		// the context wide.
		const start = Math.max(0, index - context);
		let end = index;
		let equalRun = 0;
		let cursor = index;
		while (cursor < ops.length) {
			if (ops[cursor].kind === 'equal') {
				equalRun++;
				if (equalRun > 2 * context) {
					break;
				}
			} else {
				equalRun = 0;
				end = cursor + 1;
			}
			cursor++;
		}
		const hunkEnd = Math.min(ops.length, end + context);
		const hunkBeforeStart = beforeLine - (index - start);
		const hunkAfterStart = afterLine - (index - start);
		let beforeCount = 0;
		let afterCount = 0;
		const body: string[] = [];
		for (let i = start; i < hunkEnd; i++) {
			const op = ops[i];
			if (op.kind === 'equal') {
				body.push(` ${op.line}`);
				beforeCount++;
				afterCount++;
			} else if (op.kind === 'delete') {
				body.push(`-${op.line}`);
				beforeCount++;
			} else {
				body.push(`+${op.line}`);
				afterCount++;
			}
		}
		out.push(`@@ -${range(hunkBeforeStart, beforeCount)} +${range(hunkAfterStart, afterCount)} @@`, ...body);
		// Advance the line counters over the hunk we just emitted.
		for (let i = index; i < hunkEnd; i++) {
			const op = ops[i];
			if (op.kind !== 'insert') {
				beforeLine++;
			}
			if (op.kind !== 'delete') {
				afterLine++;
			}
		}
		index = hunkEnd;
	}
	return out.join('\n');
}

function range(start: number, count: number): string {
	if (count === 1) {
		return String(start);
	}
	return `${count === 0 ? start - 1 : start},${count}`;
}
