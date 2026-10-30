// Issue #1503, measured through the VBE oracle in Excel, Word, PowerPoint and
// Access 16.0 (build 20527), with each host's own file as the referenced
// project: XLIDE reads the VBA project a reference names, and judges a name
// reached through it by what that project offers. The cases are the
// issue1503_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json; the
// engine's read is checked against a consumer and provider each host saved.

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterAll, afterEach, beforeEach, describe, expect, it } from 'vitest';
import { hostPlatform, setHostPlatform } from '../../../src/vba/hostPlatform';
import { listReferences, readModules } from '../../../src/vba/projectService';
import { referencedProjectFile } from '../../../src/vba/referencedProjectPath';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, oracleProjectReferences, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: string;
	projectReference?: { name: string; modules: OracleFixtureModule[] };
	modules: OracleFixtureModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1503_'));
const BINARIES = path.join(__dirname, '..', '..', 'fixtures', 'binaries');

function errors(oracleCase: OracleCase, read = true): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const references = oracleProjectReferences(oracleCase);
	if (!read) {
		delete references.referencedProjectSurfaces;
	}
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		...references,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`));
}

describe('names reached through a referenced VBA project (issue #1503)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(76);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			expect(found.length).toBeGreaterThan(0);
		}
	});

	it('says what the referenced project lacks', () => {
		const byId = (id: string): OracleCase => measured.find((oracleCase) => oracleCase.id === id)!;
		expect(errors(byId('issue1503_missing_member_compile'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Service.Nope'."]);
		expect(errors(byId('issue1503_private_module_project_member_compile'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Value'."]);
		expect(errors(byId('issue1503_private_class_type_compile'))[0]).toMatch(/^missing-library-reference: 'ProbeLibrary' is a VBA project this project references, and it offers no type 'Thing'/);
	});

	it('keeps an unread referenced project reported, so a refused call is not turned silent', () => {
		const unread = errors(measured.find((oracleCase) => oracleCase.id === 'issue1503_qualified_compile')!, false);
		expect(unread).toEqual([expect.stringMatching(/^undeclared-variable: Variable not defined: 'ProbeLibrary'/)]);
	});
});

describe('the referenced project, as the engine reads it (issue #1503)', () => {
	const made: string[] = [];
	afterEach(() => {
		for (const dir of made.splice(0)) {
			fs.rmSync(dir, { recursive: true, force: true });
		}
	});
	/** Excel's consumer and provider, side by side in a folder of their own. */
	const copied = (): string => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-1503-'));
		made.push(dir);
		for (const name of ['ProbeConsumerFixture.xlsm', 'ProbeLibraryFixture.xlsm']) {
			fs.copyFileSync(path.join(BINARIES, name), path.join(dir, name));
		}
		return dir;
	};

	it('reads the provider beside the consumer through the relative path, its private module hidden', () => {
		const dir = copied();
		const surfaces = readModules(path.join(dir, 'ProbeConsumerFixture.xlsm'))[0].referencedProjects;
		const probe = surfaces?.find((surface) => surface.name === 'ProbeLibrary');
		expect(probe?.modules.find((module) => module.name === 'Service')).toMatchObject({ kind: 'standard', visible: true, members: ['Value'] });
		expect(probe?.modules.find((module) => module.name === 'Secret')).toMatchObject({ visible: false, members: [] });
	});

	// The consumer also records the provider's absolute path where Excel
	// saved it, in the checkout the fixtures were made in. Where that path
	// exists, the engine rightly reads it, so this runs everywhere else.
	const absolute = /^\*\\[A-Za-z](.+)$/s.exec(listReferences(path.join(BINARIES, 'ProbeConsumerFixture.xlsm'))
		.find((reference) => reference.name === 'ProbeLibrary')?.libid ?? '')?.[1];
	it.runIf(path.sep !== String.fromCharCode(92) || !absolute || !fs.existsSync(absolute))('leaves the reference unread when the provider is gone', () => {
		const dir = copied();
		fs.rmSync(path.join(dir, 'ProbeLibraryFixture.xlsm'));
		expect(readModules(path.join(dir, 'ProbeConsumerFixture.xlsm'))[0].referencedProjects).toBeUndefined();
	});
});

// Each host's consumer and provider, saved by the host itself, copied
// together to a folder of their own, as a user moves a project and its add-in.
// The VBE finds the provider there in all four hosts (Microsoft 365, 16.0
// build 20527). Only Excel's consumer records a relative path that names it;
// Word's and PowerPoint's repeat the absolute path, and Access's is relative
// to another folder, so XLIDE finds those by the file name.
describe('a provider moved together with its consumer', () => {
	const made: string[] = [];
	const reads: string[] = [];
	// Every file outside the folders made here is hidden from the lookup, so
	// the provider where the consumer was saved, which exists on the machine
	// that saved it, is never found instead of the moved one.
	const real = hostPlatform();
	const inside = (filePath: string): boolean => made.some((dir) => {
		const relative = path.relative(dir, filePath);
		return !relative.startsWith('..') && !path.isAbsolute(relative);
	});
	beforeEach(() => {
		setHostPlatform({
			...real,
			readFile: (filePath) => {
				reads.push(path.resolve(filePath));
				return real.readFile(filePath);
			},
			statIfPresent: (filePath) => (inside(path.resolve(filePath)) ? real.statIfPresent(filePath) : undefined),
		});
	});
	afterEach(() => {
		setHostPlatform(real);
		reads.splice(0);
		for (const dir of made.splice(0)) {
			fs.rmSync(dir, { recursive: true, force: true });
		}
	});
	const hosts = [
		{ host: 'Excel', consumer: 'ProbeConsumerFixture.xlsm', provider: 'ProbeLibraryFixture.xlsm' },
		{ host: 'Word', consumer: 'ProbeConsumerFixture.docm', provider: 'ProbeLibraryFixture.dotm' },
		{ host: 'PowerPoint', consumer: 'ProbeConsumerFixture.pptm', provider: 'ProbeLibraryFixture.ppam' },
		{ host: 'Access', consumer: 'ProbeConsumerFixture.accdb', provider: 'ProbeLibraryFixture.accdb' },
	];
	/** The libids the consumer records for ProbeLibrary. */
	const recorded = (consumer: string) => listReferences(path.join(BINARIES, consumer))
		.find((reference) => reference.name === 'ProbeLibrary');
	/** The consumer in a new folder, and the provider beside it under the name its reference records when wanted. */
	const moved = (consumer: string, provider?: string): { consumer: string; provider: string } => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-moved-'));
		made.push(dir);
		const files = {
			consumer: path.resolve(dir, consumer),
			provider: path.resolve(dir, recorded(consumer)!.libid.split(/[\\/]/).pop()!),
		};
		fs.copyFileSync(path.join(BINARIES, consumer), files.consumer);
		if (provider) {
			fs.copyFileSync(path.join(BINARIES, provider), files.provider);
		}
		return files;
	};

	it.each(hosts.map((one) => [one.host, one.consumer, one.provider]))('%s: reads the provider beside the consumer, its private module hidden', (_host, consumer, provider) => {
		const files = moved(consumer, provider);
		const probe = readModules(files.consumer)[0].referencedProjects?.find((surface) => surface.name === 'ProbeLibrary');
		expect(reads).toContain(files.provider);
		expect(probe?.modules.find((module) => module.name === 'Service')).toMatchObject({ kind: 'standard', visible: true });
		expect(probe?.modules.find((module) => module.name === 'Service')?.members).toContain('Value');
		expect(probe?.modules.find((module) => module.name === 'Secret')).toMatchObject({ visible: false, members: [] });
	});

	it.each(hosts.map((one) => [one.host, one.consumer]))('%s: leaves the reference unread when the consumer moved without its provider', (_host, consumer) => {
		expect(readModules(moved(consumer).consumer)[0].referencedProjects).toBeUndefined();
	});
});

describe('which referenced files XLIDE may read (issue #1503)', () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-1503-paths-'));
	const provider = path.join(dir, 'Lib.xlam');
	fs.writeFileSync(provider, 'not empty');
	fs.writeFileSync(path.join(dir, 'Lib.txt'), 'not empty');
	const consumer = path.join(dir, 'Book.xlsm');
	afterAll(() => {
		fs.rmSync(dir, { recursive: true, force: true });
	});

	it('finds a local provider through the relative path', () => {
		expect(referencedProjectFile('*\\CZ:\\Nowhere\\Lib.xlam', '*\\CLib.xlam', consumer)).toBe(provider);
	});

	it('finds a provider beside the referring file by the name its reference records', () => {
		// As Word and PowerPoint record it, the absolute path twice.
		expect(referencedProjectFile('*\\CZ:\\Moved\\Lib.xlam', '*\\CZ:\\Moved\\Lib.xlam', consumer)).toBe(provider);
		// As Access records it, relative to another folder.
		expect(referencedProjectFile('*\\CZ:\\Moved\\Lib.xlam', '*\\C..\\..\\Moved\\Lib.xlam', consumer)).toBe(provider);
		expect(referencedProjectFile('*\\CZ:\\Moved\\Other.xlam', '*\\CZ:\\Moved\\Other.xlam', consumer)).toBeUndefined();
		// The relative path's name alone.
		expect(referencedProjectFile('*\\CZ:\\Moved\\Other.xlam', '*\\CZ:\\Moved\\Lib.xlam', consumer)).toBe(provider);
	});

	it('tries the recorded paths before the names, and the absolute name before the relative one', () => {
		const sub = path.join(dir, 'Sub');
		fs.mkdirSync(sub, { recursive: true });
		const nested = path.join(sub, 'Lib.xlam');
		fs.writeFileSync(nested, 'not empty');
		const other = path.join(dir, 'Other.xlam');
		fs.writeFileSync(other, 'not empty');
		try {
			expect(referencedProjectFile('*\\CZ:\\Moved\\Lib.xlam', '*\\CSub\\Lib.xlam', consumer)).toBe(nested);
			expect(referencedProjectFile('*\\CZ:\\Moved\\Other.xlam', '*\\CZ:\\Moved\\Lib.xlam', consumer)).toBe(other);
			if (path.sep === String.fromCharCode(92)) {
				expect(referencedProjectFile(`*\\C${nested}`, '*\\COther.xlam', consumer)).toBe(nested);
			}
		} finally {
			fs.rmSync(sub, { recursive: true, force: true });
			fs.rmSync(other, { force: true });
		}
	});

	it('takes no name that is empty, a folder, a drive or a stream', () => {
		// Where a file name may hold a colon, such files exist, so only the
		// rule keeps them out; on Windows they would name a drive or a stream.
		const colons = ['C:Lib.xlam', 'Book.xlsm:Lib.xlam'].map((name) => path.join(dir, name));
		if (path.sep === '/') {
			for (const file of colons) {
				fs.writeFileSync(file, 'not empty');
			}
		}
		try {
			for (const libid of ['*\\CZ:\\Moved\\', '*\\CZ:\\Moved\\..', '*\\CZ:', '*\\CC:Lib.xlam', '*\\CZ:\\Moved\\Book.xlsm:Lib.xlam']) {
				expect(referencedProjectFile(libid, libid, consumer)).toBeUndefined();
			}
		} finally {
			if (path.sep === '/') {
				for (const file of colons) {
					fs.rmSync(file, { force: true });
				}
			}
		}
	});

	it('never reads a network or device path', () => {
		expect(referencedProjectFile('*\\C\\\\host\\share\\Lib.xlam', undefined, consumer)).toBeUndefined();
		expect(referencedProjectFile('*\\C\\\\?\\C:\\Lib.xlam', undefined, consumer)).toBeUndefined();
		expect(referencedProjectFile('*\\CZ:\\x.xlam', '*\\C\\\\host\\share\\Lib.xlam', consumer)).toBeUndefined();
		expect(referencedProjectFile('*\\CZ:\\x.xlam', '*\\CLib.xlam', '\\\\host\\share\\Book.xlsm')).toBeUndefined();
	});

	it('reads only an Office file a reference can name, and never the referring file itself', () => {
		expect(referencedProjectFile('*\\CZ:\\x.xlam', '*\\CLib.txt', consumer)).toBeUndefined();
		fs.writeFileSync(consumer, 'not empty');
		expect(referencedProjectFile('*\\CZ:\\x.xlam', '*\\CBook.xlsm', consumer)).toBeUndefined();
	});
});
