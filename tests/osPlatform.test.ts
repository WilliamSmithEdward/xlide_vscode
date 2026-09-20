import { afterEach, describe, expect, it } from 'vitest';
import { isWindows, osPlatform } from '../src/util/osPlatform';

// osPlatform() has to READ process.platform on every call, not capture it at
// import time. `process.platform` never changes in production, so a constant
// would be correct there - but the Office coordination tests redefine it to
// exercise Windows-only code on any machine, and a constant freezes before
// they can. That is not hypothetical: shipping it as a constant left those
// tests green on Windows and red on CI's Linux runner, because on Windows the
// captured value happened to be the one the test wanted.
//
// This guard fails on every platform, which is the point.

const real = Object.getOwnPropertyDescriptor(process, 'platform')!;

afterEach(() => {
	Object.defineProperty(process, 'platform', real);
});

function pretend(platform: string): void {
	Object.defineProperty(process, 'platform', { ...real, value: platform });
}

describe('osPlatform', () => {
	it('follows a redefined process.platform', () => {
		// Two values, so this cannot pass by coincidence on the host platform.
		pretend('win32');
		expect(osPlatform()).toBe('win32');

		pretend('linux');
		expect(osPlatform()).toBe('linux');

		pretend('darwin');
		expect(osPlatform()).toBe('darwin');
	});

	it('reports the real platform once it is restored', () => {
		pretend('aix');
		expect(osPlatform()).toBe('aix');

		Object.defineProperty(process, 'platform', real);
		expect(osPlatform()).toBe(process.platform);
	});

	it('isWindows follows it too', () => {
		pretend('win32');
		expect(isWindows()).toBe(true);

		pretend('linux');
		expect(isWindows()).toBe(false);
	});
});
