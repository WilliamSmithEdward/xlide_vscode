// Diagnostics tests: Let and Set into array elements, Type fields and
// Collection or Dictionary items (issue #306). Each sample was run or
// compiled through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeProjectModule } from './helpers';

const DICT = 'Dim d As Object\n    Set d = CreateObject("Scripting.Dictionary")\n    ';

function module(body: string, head = ''): string {
	return `Option Explicit\n${head}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
}

function found(body: string, head = ''): string[] {
	return analyzeModule(module(body, head)).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code} ${/Run-time error '(\d+)'|compile error/i.exec(diag.message)?.[0] ?? ''}`.trim());
}

describe('an element or item assigned the wrong way (issue #306)', () => {
	it('is a compile error into an array element', () => {
		expect(found('Dim v(1) As Variant\n    Set v(0) = 5')).toEqual(['set-requires-object compile error']);
		expect(found('Dim v() As Variant\n    ReDim v(1)\n    Set v(0) = 5')).toEqual(['set-requires-object compile error']);
		expect(found('Dim a(1) As Collection\n    a(0) = New Collection')).toEqual(['set-required compile error']);
	});

	it('raises 450 on a Let of New Collection into a Variant element, field or item', () => {
		expect(found('Dim v(1) As Variant\n    v(0) = New Collection')).toEqual(["object-default-value Run-time error '450'"]);
		expect(found('Dim v() As Variant\n    ReDim v(1)\n    v(0) = New Collection')).toEqual(["object-default-value Run-time error '450'"]);
		expect(found('Dim t As T1\n    t.v = New Collection', 'Private Type T1\n    v As Variant\nEnd Type\n')).toEqual(["object-default-value Run-time error '450'"]);
		expect(found(`${DICT}d("k") = New Collection`)).toEqual(["object-default-value Run-time error '450'"]);
		expect(found('Dim c As New Collection\n    c.Add New Collection\n    Dim x As Variant\n    x = c(1)')).toEqual(["object-default-value Run-time error '450'"]);
	});

	it('raises 438 or 424 on a Set into a Collection item, by what it holds', () => {
		expect(found('Dim c As New Collection\n    c.Add New Collection\n    Set c(1) = New Collection')).toEqual(["runtime-member-not-found Run-time error '438'"]);
		expect(found('Dim c As New Collection\n    c.Add New Collection\n    Set c.Item(1) = New Collection')).toEqual(["runtime-member-not-found Run-time error '438'"]);
		expect(found('Dim c As New Collection\n    c.Add 5, "k"\n    Set c("k") = New Collection')).toEqual(["variant-value-misuse Run-time error '424'"]);
	});

	it('raises 424 on a Set from an item that holds a value', () => {
		expect(found('Dim c As New Collection\n    c.Add 5\n    Dim x As Object\n    Set x = c(1)')).toEqual(["variant-value-misuse Run-time error '424'"]);
		expect(found('Dim c As New Collection\n    c.Add "s"\n    Dim x As Object\n    Set x = c.Item(1)')).toEqual(["variant-value-misuse Run-time error '424'"]);
		expect(found(`${DICT}d("k") = 5\n    Dim x As Object\n    Set x = d("k")`)).toEqual(["variant-value-misuse Run-time error '424'"]);
		expect(found(`${DICT}d.Add "k", 5\n    Dim x As Object\n    Set x = d("k")`)).toEqual(["variant-value-misuse Run-time error '424'"]);
	});

	it('raises 13 on ActiveSheet Set into a Collection or a class', () => {
		expect(found('Dim a(1) As Collection\n    Set a(0) = ActiveSheet')).toEqual(["assignment-object-type-mismatch Run-time error '13'"]);
		expect(found('Dim c As Collection\n    Set c = ActiveSheet')).toEqual(["assignment-object-type-mismatch Run-time error '13'"]);
		const caller = module('Dim k As Class1\n    Set k = ActiveSheet');
		const diagnostics = analyzeProjectModule(caller, [{ moduleName: 'Class1', moduleKind: 'class', source: 'Option Explicit\nPublic V As Long\n' }], 'Module1');
		expect(diagnostics.filter((diag) => diag.code === 'assignment-object-type-mismatch')).toHaveLength(1);
	});
});

describe('an element or item assigned as its type allows (issue #306)', () => {
	it('stays quiet', () => {
		const bodies = [
			`${DICT}d("k") = 5\n    Main = d("k")`,
			`${DICT}Set d("k") = New Collection\n    Main = d("k").Count`,
			`${DICT}d("k") = 5\n    Set d("k") = New Collection\n    Dim x As Object\n    Set x = d("k")\n    Main = x.Count`,
			'Dim v(1) As Variant\n    Set v(0) = New Collection\n    Main = v(0).Count',
			'Dim a(1) As Collection\n    Set a(0) = New Collection\n    Main = a(0).Count',
			'Dim c As New Collection\n    c.Add New Collection\n    c(1).Add 5\n    Main = c(1).Count',
			'Dim c As New Collection\n    c.Add 5\n    Dim x As Variant\n    x = c(1)\n    Main = x',
			'Dim c As New Collection\n    c.Add New Collection\n    Dim x As Object\n    Set x = c(1)\n    Main = x.Count',
			'Dim v As Variant\n    v = Array(1, 2)\n    Set v(0) = New Collection',
			'Dim a(1) As Object\n    Set a(0) = ActiveSheet\n    Main = a(0).Name',
			'Dim a(1) As Range\n    Set a(0) = Range("A1")\n    a(0) = 5\n    Main = Range("A1").Value',
			'Range("A1") = 5\n    Main = Range("A1").Value',
		];
		for (const body of bodies) {
			expect(found(body), body).toEqual([]);
		}
	});
});
