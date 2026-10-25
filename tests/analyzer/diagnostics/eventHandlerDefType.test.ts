// An untyped event parameter is the type its own module's DefType lines give
// it, on the event's side and on the handler's. Measured in Excel, Word,
// PowerPoint and Access 16.0 (build 20430) through the VBE oracle: the cases
// are event_deftype_* in syntax_corpus/oracle/vbe_oracle_cases.json, one per
// host, and the four hosts agree on every one.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { AnalysisWorkerState } from '../../../src/analysisWorkerLogic';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: 'excel' | 'word' | 'powerpoint' | 'access';
	modules: OracleFixtureModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const HOSTS = ['', 'word_', 'powerpoint_', 'access_'] as const;

function caseFor(id: string): OracleCase {
	const found = corpus.cases.find((oracleCase) => oracleCase.id === id);
	if (!found) {
		throw new Error(`no oracle case ${id}`);
	}
	return found;
}

/** An Access form's module reaches the analyzer as the extension reads it from the database. */
function moduleInput(oracleCase: OracleCase, mod: OracleFixtureModule): VbaProjectModuleInput {
	const input = oracleModuleInput(mod);
	return oracleCase.host === 'access' && mod.type === 'document' ? { ...input, type: 'accessform' } : input;
}

/** Each error as `code: spanned text`, analyzed as the extension analyzes the case's host. */
function errorsFor(oracleCase: OracleCase): string[] {
	const modules = oracleCase.modules.map((mod) => moduleInput(oracleCase, mod));
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: oracleCase.host ?? 'excel',
		...(mod.designerClass ? { designerClass: mod.designerClass } : {}),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics
		.filter((diagnostic) => diagnostic.severity === 'error')
		.map((diagnostic) => `${diagnostic.code}: ${mod.source.slice(diagnostic.span.start, diagnostic.span.end)}`));
}

function eachHost(name: string): Array<[string]> {
	return HOSTS.map((host) => [`event_deftype_${host}${name}_compile`]);
}

describe("an untyped event parameter takes its own module's DefType", () => {
	it.each([...eachHost('mismatch'), ...eachHost('publisher_variant')])('%s refuses the handler', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual(['event-handler-signature: source_Pulse']);
	});

	it.each([
		...eachHost('matching'),
		...eachHost('override'),
		...eachHost('mismatch_empty'),
		...eachHost('suffix'),
	])('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});

	it('says where each side took its type from', () => {
		const messages = sinkMessages(caseFor('event_deftype_mismatch_compile'));
		expect(messages).toHaveLength(1);
		expect(messages[0]).toContain("'x' is String (from this module's Deftype), and the event's is Long (from its module's Deftype)");
	});
});

describe("an untyped handler parameter of a library event takes its module's DefType", () => {
	const HANDLERS: Record<string, string> = {
		'': 'app_WorkbookBeforeClose',
		word_: 'app_DocumentBeforeClose',
		powerpoint_: 'app_PresentationBeforeSave',
		access_: 'Form_BeforeUpdate',
	};

	it.each(HOSTS.map((host) => [`event_deftype_${host}library_mismatch_compile`, HANDLERS[host]]))('%s refuses %s', (id, handler) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual([`event-handler-signature: ${handler}`]);
	});

	it.each(eachHost('library_matching'))('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});
});

describe('DefType reaches every untyped handler parameter, and no other', () => {
	it.each([...eachHost('byval_mismatch'), ...eachHost('array_mismatch')])('%s refuses the handler', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual(['event-handler-signature: source_Pulse']);
	});

	it.each([...eachHost('handler_suffix'), ...eachHost('array_matching'), ...eachHost('userform_matching')])('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});

	it.each(eachHost('userform_mismatch'))('%s refuses UserForm_QueryClose', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual(['event-handler-signature: UserForm_QueryClose']);
	});

	// PowerPoint has no document module, and an Access form's is above.
	it.each([
		['event_deftype_document_matching_compile', []],
		['event_deftype_document_mismatch_compile', ['event-handler-signature: Workbook_BeforeClose']],
		['event_deftype_word_document_matching_compile', []],
		['event_deftype_word_document_mismatch_compile', ['event-handler-signature: Document_ContentControlOnExit']],
	])('%s: a document module handler', (id, errors) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe(errors.length > 0 ? 'rejected' : 'accepted');
		expect(errorsFor(oracleCase)).toEqual(errors);
	});
});

describe('a DefType edit in the class that declares the event reaches its handlers', () => {
	const SINK = 'Option Explicit\nDefLng X\nPrivate WithEvents source As Publisher\nPrivate Sub source_Pulse(ByRef x)\n    Debug.Print x\nEnd Sub\n';

	it('re-derives the handler module in the analysis worker', () => {
		const state = new AnalysisWorkerState();
		const findings = (deftype: string, generation: number): number => {
			state.handle({
				kind: 'seed',
				projectKey: 'wb',
				generation,
				modules: [
					{ moduleName: 'Publisher', source: `Option Explicit\n${deftype}\nPublic Event Pulse(ByRef x)\n`, type: 'class' },
					{ moduleName: 'Sink', source: SINK, type: 'class' },
				],
			});
			const response = state.handle({
				kind: 'analyze',
				requestId: generation,
				docKey: 'doc:Sink',
				projectKey: 'wb',
				generation,
				moduleName: 'Sink',
				source: SINK,
				moduleType: 'class',
				moduleKind: 'class',
			});
			return (response?.kind === 'result' ? response.diagnostics : []).filter((d) => d.code === 'event-handler-signature').length;
		};
		expect(findings('DefLng X', 1)).toBe(0);
		expect(findings('DefStr X', 2)).toBe(1);
		expect(findings('DefLng X', 3)).toBe(0);
	});
});

/** The event-handler-signature messages for a case's Sink module. */
function sinkMessages(oracleCase: OracleCase): string[] {
	const modules = oracleCase.modules.map((mod) => moduleInput(oracleCase, mod));
	const project = buildVbaProjectIndex(modules);
	return analyzeVbaModuleSource({
		source: modules.find((mod) => mod.moduleName === 'Sink')!.source,
		moduleName: 'Sink',
		moduleKind: 'class',
		...projectAnalysisOptionsForModule(project, 'Sink', projectProcedureSignatures(project)),
	}).diagnostics.filter((diagnostic) => diagnostic.code === 'event-handler-signature').map((diagnostic) => diagnostic.message);
}
