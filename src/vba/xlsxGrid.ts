// Where a worksheet's cells are, in EMU: what a drawing's cell anchors are
// measured against.
//
// A drawing anchor names a cell and an offset into it, so reading or moving a
// shape by cells needs no geometry. Rotating one does: Excel keeps the anchor
// of a shape turned between 45 and 135 degrees (or 225 and 315) as its box
// turned a quarter about its center, measured by having Excel rotate a group
// to 45 degrees and save. Swapping the box needs the column widths and row
// heights around it.
//
// Excel's own grid follows the screen: a column's width in points is its
// width in pixels at the display's DPI, so one file lays out a few points
// differently from machine to machine (measured at 144 dpi, where five
// characters are 31 points; at 96 dpi they are 30). The grid here is Excel's
// at 96 dpi, from the conversions the file format defines, so a swapped
// anchor lands where Excel at 100% scaling would put it and within a pixel
// or two of it anywhere else.

import { attr, findElement } from './ooxml';

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;

/**
 * The widest digit of the workbook's Normal font, in pixels at 96 dpi, which
 * column widths are counted in. 7 for Calibri 11, Aptos Narrow 11 and Arial
 * 10, the defaults of every Excel version, which measured alike; a workbook
 * whose Normal font is unusual gets the same, and a rotated shape in it can
 * land a little off.
 */
const MAX_DIGIT_WIDTH = 7;

/** A position in a sheet as a drawing marker gives it: a cell, and EMU into it. */
export interface GridMarker {
	col: number;
	colOff: number;
	row: number;
	rowOff: number;
}

export interface SheetGrid {
	/** The left edge of a 0-based column, in EMU. */
	columnLeft(col: number): number;
	/** The top edge of a 0-based row, in EMU. */
	rowTop(row: number): number;
	/** The marker at a position in EMU. */
	marker(x: number, y: number): GridMarker;
	/** A marker's position in EMU. */
	position(marker: GridMarker): { x: number; y: number };
}

/** A column width in characters as pixels, by the conversion the file format defines. */
function widthToPixels(width: number): number {
	return Math.trunc(((256 * width + Math.trunc(128 / MAX_DIGIT_WIDTH)) / 256) * MAX_DIGIT_WIDTH);
}

/** The grid of a worksheet part. */
export function sheetGrid(sheetXml: string): SheetGrid {
	const format = /<sheetFormatPr\b[^>]*>/.exec(sheetXml)?.[0] ?? '';
	const defaultColWidth = attr(format, 'defaultColWidth');
	const baseColWidth = Number(attr(format, 'baseColWidth') ?? 8);
	// Excel's default column is the base width in digits plus five pixels of
	// padding, rounded up to a multiple of eight: 64 pixels for base 8.
	const defaultColumnPixels = defaultColWidth !== undefined
		? widthToPixels(Number(defaultColWidth))
		: Math.ceil((baseColWidth * MAX_DIGIT_WIDTH + 5) / 8) * 8;
	const defaultRowPoints = Number(attr(format, 'defaultRowHeight') ?? 15);

	const columns = new Map<number, number>();
	const cols = findElement(sheetXml, 'cols');
	if (cols) {
		for (const [tag] of sheetXml.slice(cols.openEnd, cols.end).matchAll(/<col\b[^>]*>/g)) {
			const min = Number(attr(tag, 'min') ?? 0);
			const max = Math.min(Number(attr(tag, 'max') ?? min), 16384);
			const hidden = attr(tag, 'hidden') === '1' || attr(tag, 'hidden') === 'true';
			const width = attr(tag, 'width');
			const pixels = hidden ? 0 : width !== undefined ? widthToPixels(Number(width)) : defaultColumnPixels;
			for (let c = min; c <= max; c++) { columns.set(c - 1, pixels); }
		}
	}
	const rows = new Map<number, number>();
	for (const [tag] of sheetXml.matchAll(/<row\b[^>]*>/g)) {
		const r = Number(attr(tag, 'r') ?? 0);
		if (!r) { continue; }
		const hidden = attr(tag, 'hidden') === '1' || attr(tag, 'hidden') === 'true';
		const ht = attr(tag, 'ht');
		if (hidden) { rows.set(r - 1, 0); } else if (ht !== undefined) { rows.set(r - 1, Number(ht)); }
	}

	const columnWidth = (col: number): number => (columns.get(col) ?? defaultColumnPixels) * EMU_PER_PIXEL;
	const rowHeight = (row: number): number => Math.round((rows.get(row) ?? defaultRowPoints) * EMU_PER_POINT);
	const columnLeft = (col: number): number => {
		let x = 0;
		for (let c = 0; c < col; c++) { x += columnWidth(c); }
		return x;
	};
	const rowTop = (row: number): number => {
		let y = 0;
		for (let r = 0; r < row; r++) { y += rowHeight(r); }
		return y;
	};
	return {
		columnLeft,
		rowTop,
		marker(x, y) {
			let col = 0;
			let left = 0;
			while (col < 16383 && left + columnWidth(col) <= Math.max(0, x)) { left += columnWidth(col); col++; }
			let row = 0;
			let top = 0;
			while (row < 1048575 && top + rowHeight(row) <= Math.max(0, y)) { top += rowHeight(row); row++; }
			return { col, colOff: Math.max(0, Math.round(x - left)), row, rowOff: Math.max(0, Math.round(y - top)) };
		},
		position(marker) {
			return { x: columnLeft(marker.col) + marker.colOff, y: rowTop(marker.row) + marker.rowOff };
		},
	};
}
