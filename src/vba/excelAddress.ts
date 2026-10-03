// Pure worksheet address primitives shared by file handling and VBA analysis.

/** The last row and column of a worksheet: XFD1048576. */
export const MAX_ROW = 1048576;
export const MAX_COLUMN = 16384;

export function columnToIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) {
		n = n * 26 + (ch.charCodeAt(0) - 64);
	}
	return n;
}

export function indexToColumn(index: number): string {
	let n = index;
	let out = '';
	while (n > 0) {
		const rem = (n - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		n = Math.floor((n - 1) / 26);
	}
	return out;
}
