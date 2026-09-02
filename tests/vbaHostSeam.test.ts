import { describe, expect, it } from 'vitest';
import {
	EMPTY_HOST_MODEL,
	hostObjectModelForToken,
	hostTokenForFileName,
	registerHostObjectModel,
} from '../src/analyzer/host/hostRegistry';
import { getExcelObjectModel } from '../src/analyzer/host/excelObjectModel';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';

// Issue #24. The resolvers have always accepted any HostObjectModel and
// defaulted to Excel; this is the seam that lets a caller choose, with the
// semantics the issue asked for: absent means Excel and nothing changes, and
// a NAMED host with no model yet asserts no host knowledge at all - because
// telling Word's ThisDocument it has Cells and Range was the original bug.

describe('the host registry', () => {
	it('answers nothing for absent or excel, so defaults ride', () => {
		expect(hostObjectModelForToken(undefined)).toBeUndefined();
		expect(hostObjectModelForToken('excel')).toBeUndefined();
		expect(hostObjectModelForToken('  Excel ')).toBeUndefined();
	});

	it('answers the empty model for a named host with no model yet', () => {
		for (const token of ['outlook', 'visio', 'project', 'other']) {
			const model = hostObjectModelForToken(token);
			expect(model, token).toBe(EMPTY_HOST_MODEL);
			expect(Object.keys(model!.types)).toEqual([]);
			expect(Object.keys(model!.globals)).toEqual([]);
		}
	});

	it('answers a registered model for its token', () => {
		registerHostObjectModel('project', () => getExcelObjectModel());
		try {
			expect(hostObjectModelForToken('project')).toBe(getExcelObjectModel());
		} finally {
			// Put the throwaway registration back to "no model".
			registerHostObjectModel('project', () => EMPTY_HOST_MODEL);
		}
	});

	it('derives the host from the container file name', () => {
		expect(hostTokenForFileName('Book.xlsm')).toBe('excel');
		expect(hostTokenForFileName('Book.XLSB')).toBe('excel');
		expect(hostTokenForFileName('Report.docm')).toBe('word');
		expect(hostTokenForFileName('Deck.pptm')).toBe('powerpoint');
		expect(hostTokenForFileName('Data.accdb')).toBe('access');
		expect(hostTokenForFileName('Notes.txt')).toBeUndefined();
	});
});

describe('analysis under a named host', () => {
	// `Volatile` is an Excel Application member injected into the bare global
	// scope, so under Excel a bare call to it is known; under Word it is not
	// an assertion anyone can make.
	const SOURCE = [
		'Option Explicit',
		'Public Sub Recalc()',
		'    Volatile',
		'End Sub',
		'',
	].join('\r\n');

	function unknownCalls(host: string | undefined): string[] {
		return analyzeVbaModuleSource({
			source: SOURCE,
			moduleName: 'Module1',
			moduleType: 'standard',
			host,
			knownProcedures: new Set(['recalc']),
		} as never).diagnostics
			.filter((d) => d.code === 'unknown-call')
			.map((d) => d.message);
	}

	it('absent host keeps Excel behavior: Volatile is a known bare call', () => {
		expect(unknownCalls(undefined)).toEqual([]);
	});

	it('excel says the same thing explicitly', () => {
		expect(unknownCalls('excel')).toEqual([]);
	});

	it('word does not pretend to know Excel Application members', () => {
		expect(unknownCalls('word').some((m) => m.includes('Volatile'))).toBe(true);
	});
});

describe('the worker carries the host token', () => {
	const SOURCE = [
		'Option Explicit',
		'Public Sub Recalc()',
		'    Volatile',
		'End Sub',
		'',
	].join('\r\n');

	it('analyzes under the requested host and re-analyzes when it changes', () => {
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', projectKey: 'wb-host', generation: 1,
			modules: [{ moduleName: 'Module1', source: SOURCE, type: 'standard' }],
		});
		const excel = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-host', projectKey: 'wb-host', generation: 1,
			source: SOURCE, moduleName: 'Module1', moduleType: 'standard',
		});
		expect(excel?.kind === 'result'
			? excel.diagnostics.filter((d) => d.code === 'unknown-call')
			: undefined).toEqual([]);

		// Same document, host changed: the fingerprint must not reuse the
		// Excel answer.
		const word = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc-host', projectKey: 'wb-host', generation: 1,
			source: SOURCE, moduleName: 'Module1', moduleType: 'standard', host: 'word',
		});
		expect(word?.kind).toBe('result');
		if (word?.kind !== 'result') { return; }
		expect(word.diagnostics.some((d) => d.code === 'unknown-call' && d.message.includes('Volatile')))
			.toBe(true);
	});
});

describe('host constants stay inside their host (no cross-host leakage)', () => {
	const diagnostics = (source: string, host?: string) =>
		analyzeVbaModuleSource({ source, moduleName: 'Module1', moduleKind: 'standard', host })
			.diagnostics;

	it('folds xlLandscape in Excel, never in Word', () => {
		// xlLandscape = 2, so the divisor folds to zero under Excel's model.
		const source = [
			'Sub T()',
			'    Dim x As Double',
			'    x = 1 / (xlLandscape - 2)',
			'End Sub',
			'',
		].join('\r\n');
		expect(diagnostics(source).some((d) => d.code === 'division-by-zero')).toBe(true);
		expect(diagnostics(source, 'word').some((d) => d.code === 'division-by-zero')).toBe(false);
	});

	it('folds wdMainTextStory in Word, never in Excel', () => {
		// wdMainTextStory = 1, so the divisor folds to zero under Word's model.
		const source = [
			'Sub T()',
			'    Dim x As Double',
			'    x = 1 / (wdMainTextStory - 1)',
			'End Sub',
			'',
		].join('\r\n');
		expect(diagnostics(source, 'word').some((d) => d.code === 'division-by-zero')).toBe(true);
		expect(diagnostics(source).some((d) => d.code === 'division-by-zero')).toBe(false);
	});

	it('resolves the Word-qualified form under Word only', () => {
		const source = [
			'Sub T()',
			'    Dim x As Double',
			'    x = 1 / (Word.wdMainTextStory - 1)',
			'End Sub',
			'',
		].join('\r\n');
		expect(diagnostics(source, 'word').some((d) => d.code === 'division-by-zero')).toBe(true);
		expect(diagnostics(source).some((d) => d.code === 'division-by-zero')).toBe(false);
	});
});
