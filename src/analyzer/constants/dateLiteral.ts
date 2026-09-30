export const DAY_MS = 86400000;
/** Serial 0, December 30, 1899, as a UTC time. */
export const DATE_EPOCH_MS = Date.UTC(1899, 11, 30);

/**
 * A date literal's serial: the days from December 30, 1899, with the time of
 * day as the fraction. `#12/31/9999#` is 2958465, `#1/1/100#` is -657434
 * (issue #203), `#12:00:00 PM#` is 0.5 and `#1/1/2000 6:00 AM#` is 36526.25
 * (issue #208). Read are `#m/d/yyyy#` and `#yyyy-mm-dd#`, a time of
 * `h:mm[:ss]` with an optional AM or PM, and the two together; a month name
 * or a two-digit year is left to the VBE.
 */
export function dateLiteralSerial(raw: string): number | undefined {
	const text = raw.replace(/^#|#$/g, '').trim();
	const match = /^(?:(\d{1,2})\/(\d{1,2})\/(\d{3,4})|(\d{3,4})-(\d{1,2})-(\d{1,2}))?\s*(?:(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/
		.exec(text);
	if (!match || text.length === 0) {
		return undefined;
	}
	const [, usMonth, usDay, usYear, isoYear, isoMonth, isoDay, hours, minutes, seconds, meridiem] = match;
	let serial = 0;
	if (usYear !== undefined || isoYear !== undefined) {
		const year = Number(usYear ?? isoYear);
		const month = Number(usMonth ?? isoMonth);
		const day = Number(usDay ?? isoDay);
		if (year < 100 || month < 1 || month > 12 || day < 1) {
			return undefined;
		}
		const at = new Date(0);
		at.setUTCFullYear(year, month - 1, day);
		if (at.getUTCMonth() !== month - 1) {
			return undefined; // #2/30/2020# is no date
		}
		serial = Math.round((at.getTime() - DATE_EPOCH_MS) / DAY_MS);
	}
	if (hours !== undefined) {
		let hour = Number(hours);
		const minute = Number(minutes);
		const second = Number(seconds ?? 0);
		if (meridiem !== undefined) {
			if (hour < 1 || hour > 12) {
				return undefined;
			}
			hour = hour % 12 + (/^[Pp]/.test(meridiem) ? 12 : 0);
		}
		if (hour > 23 || minute > 59 || second > 59) {
			return undefined;
		}
		const fraction = (hour * 3600 + minute * 60 + second) / 86400;
		// Before 1899-12-30 the time runs the other way: -1.25 is 12/29/1899 6 AM.
		serial = serial < 0 ? serial - fraction : serial + fraction;
	}
	return serial;
}
