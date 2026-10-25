// `WithEvents x As Collection` names a project class, form or document module
// called Collection ahead of VBA's Collection, and it sources events when that
// module declares one. A Type, an Enum or a standard module of the name is no
// such module, and the declaration is refused. Measured in Excel, Word,
// PowerPoint and Access 16.0 (build 20430) through the VBE oracle; the
// collection_class_* fixtures in vbe_oracle_cases.json hold every verdict,
// the four hosts agree on each, and this test holds the analyzer to them.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';
import { analyzeProjectModule } from './helpers';

interface Fixture {
	id: string;
	host?: string;
	expected: 'accepted' | 'rejected';
	modules: OracleFixtureModule[];
}

const corpus = (JSON.parse(readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8')) as { cases: Fixture[] }).cases;
const fixtures: Fixture[] = corpus.filter((kase) => kase.id.startsWith('collection_class_'));
const vbaFixtures: Fixture[] = corpus.filter((kase) => kase.id.startsWith('collection_vba_'));

/** Each error as `code: spanned text`, analyzed as the extension analyzes the fixture's host. */
function errorsFor(fixture: Fixture): string[] {
	const modules = fixture.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules);
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: fixture.host ?? 'excel',
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics
		.filter((diagnostic) => diagnostic.severity === 'error')
		.map((diagnostic) => `${diagnostic.code}: ${mod.source.slice(diagnostic.span.start, diagnostic.span.end)}`));
}

/** What the analyzer reports for each measured verdict. */
function expectedErrors(fixture: Fixture): string[] {
	if (fixture.expected === 'accepted') {
		return [];
	}
	// The event is found in the project class, and its handler is checked against it.
	return fixture.id.endsWith('_handler_mismatch_compile')
		? ['event-handler-signature: source_Pulse']
		: ['withevents-declaration: source'];
}

describe('WithEvents As Collection against the measured fixtures', () => {
	it('has every shape in every host', () => {
		const shapes = (host: string) => new Set(fixtures
			.filter((fixture) => (fixture.host ?? 'excel') === host)
			.map((fixture) => fixture.id.replace(/^collection_class_(word_|powerpoint_|access_)?/, '')));
		// A renamed workbook or document module is Excel and Word only, a worksheet Excel only.
		const common = ['local_events', 'renamed_events', 'local_plain', 'local_events_handler', 'local_events_handler_mismatch',
			'standard_module', 'user_type', 'userform_events', 'userform_plain', 'inactive_event', 'enum'].map((shape) => `${shape}_compile`);
		const documents = ['document_events_compile', 'document_plain_compile'];
		expect(shapes('excel')).toEqual(new Set([...common, ...documents, 'sheet_events_compile', 'sheet_plain_compile']));
		expect(shapes('word')).toEqual(new Set([...common, ...documents]));
		expect(shapes('powerpoint')).toEqual(new Set(common));
		expect(shapes('access')).toEqual(new Set(common));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		expect(errorsFor(fixture)).toEqual(expectedErrors(fixture));
	});
});

// `VBA.Collection` is VBA's Collection whatever the project calls
// Collection, so it sources no events, even beside a project class or
// UserForm of the name that declares one. The collection_vba_* fixtures
// measure it in the four hosts: in any case, bracketed, or with spaces or a
// line continuation around the dot it is refused, and a plain
// `As VBA.Collection` variable beside `WithEvents source As Collection`
// compiles.
describe('WithEvents As VBA.Collection against the measured fixtures', () => {
	it('has every shape in every host', () => {
		const shapes = ['local_events', 'intrinsic_only', 'local_plain', 'userform_events', 'lowercase', 'bracketed',
			'bracketed_member', 'spaced', 'continued_before_dot', 'continued_after_dot', 'continued_twice', 'handler', 'plain_variable'].map((shape) => `${shape}_compile`);
		for (const host of ['excel', 'word', 'powerpoint', 'access']) {
			expect(new Set(vbaFixtures
				.filter((fixture) => (fixture.host ?? 'excel') === host)
				.map((fixture) => fixture.id.replace(/^collection_vba_(word_|powerpoint_|access_)?/, '')))).toEqual(new Set(shapes));
		}
	});

	it.each(vbaFixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		expect(errorsFor(fixture)).toEqual(fixture.expected === 'accepted' ? [] : ['withevents-declaration: source']);
	});

	const messages = (sink: string): string[] => analyzeProjectModule(sink, [{ moduleName: 'Sink', source: sink, type: 'class' }], 'Sink', { moduleKind: 'class' })
		.filter((diagnostic) => diagnostic.code === 'withevents-declaration')
		.map((diagnostic) => diagnostic.message);

	it('quotes a continued type on one line', () => {
		expect(messages('Option Explicit\nPrivate WithEvents source As VBA _\n    .Collection\n'))
			.toEqual(["WithEvents variable 'source' is declared As VBA .Collection, which does not source automation events."]);
	});

	it('joins a run of continuations, a bare CR terminator and DBCS whitespace as the lexer does', () => {
		expect(messages('Option Explicit\nPrivate WithEvents source As VBA _\n _\r\n .Collection\n'))
			.toEqual(["WithEvents variable 'source' is declared As VBA .Collection, which does not source automation events."]);
		expect(messages('Option Explicit\rPrivate WithEvents source As VBA. _\rCollection\r'))
			.toEqual(["WithEvents variable 'source' is declared As VBA. Collection, which does not source automation events."]);
		expect(messages('Option Explicit\nPrivate WithEvents source As VBA\u3000_\r\n\u3000.Collection\n'))
			.toEqual(["WithEvents variable 'source' is declared As VBA .Collection, which does not source automation events."]);
	});

	// Around the dot, whitespace is what the lexer takes for it: the eom
	// character is, a no-break space is not (issue #132).
	it('takes whitespace around the dot as the lexer does', () => {
		expect(messages('Option Explicit\nPrivate WithEvents source As VBA\u0019.\u3000Collection\n')).toHaveLength(1);
		expect(messages('Option Explicit\nPrivate WithEvents source As VBA\u00A0.Collection\n')).toEqual([]);
	});

	// One bracketed name is no qualified name: the VBE refuses
	// `As [VBA.Collection]` with "User-defined type not defined" in the four
	// hosts (measured, build 20430), so this rule does not claim it.
	it('leaves a single bracketed [VBA.Collection] alone', () => {
		expect(messages('Option Explicit\nPrivate WithEvents source As [VBA.Collection]\n')).toEqual([]);
	});
});

describe('WithEvents As Collection without a project module of the name', () => {
	const SINK = 'Option Explicit\nPrivate WithEvents source As Collection\n';

	function withEvents(others: Array<{ moduleName: string; source: string; type?: 'class' | 'standard' }>): string[] {
		return analyzeProjectModule(SINK, [{ moduleName: 'Sink', source: SINK, type: 'class' }, ...others], 'Sink', { moduleKind: 'class' })
			.filter((diagnostic) => diagnostic.code === 'withevents-declaration')
			.map((diagnostic) => diagnostic.message);
	}

	it("is VBA's Collection, which sources no events (issue #124)", () => {
		expect(withEvents([])).toEqual([expect.stringContaining('does not source automation events')]);
	});

	it('is not taken from an event class of another name', () => {
		expect(withEvents([{ moduleName: 'Collections', source: 'Option Explicit\nPublic Event Pulse()\n', type: 'class' }]))
			.toEqual([expect.stringContaining('does not source automation events')]);
	});

	// A standard module is no type ("A module is not a valid type", measured
	// for one with no Event); an Event declared in it, itself an error, does
	// not make it one.
	it('is not taken from a standard module named Collection that declares an Event', () => {
		expect(withEvents([{ moduleName: 'Collection', source: 'Option Explicit\nPublic Event Pulse()\n', type: 'standard' }]))
			.toEqual([expect.stringContaining('does not source automation events')]);
	});
});
