// Diagnostics tests: DAO objects read as values in Access (issue #464).
// Measured in Access 16.0 64-bit (2026-10-02) through pyVBAharness: a
// Recordset's, Database's and TableDef's default member holds a collection,
// Fields' and TableDefs' default member Item needs an index, and read whole
// each raises an error DAO chooses. A never-set Field raises 91.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const SET: Readonly<Record<string, string>> = {
	Recordset: 'Set x = db.OpenRecordset("SELECT 1 AS a")',
	Database: 'Set x = db',
	TableDef: 'Set x = db.TableDefs(0)',
	Fields: 'Set x = db.TableDefs(0).Fields',
	TableDefs: 'Set x = db.TableDefs',
	Field: 'Set x = db.OpenRecordset("SELECT 1 AS a").Fields(0)',
};

function found(type: string, set: boolean, use: string): string[] {
	const lines = ['Dim db As DAO.Database', 'Set db = CurrentDb', `Dim x As DAO.${type}`, ...(set ? [SET[type]] : []), use];
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src, { host: 'access' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a DAO object whose default member holds a collection', () => {
	it('is a Type mismatch as an operand and refuses a Let, set or not', () => {
		for (const type of ['Recordset', 'Database', 'TableDef']) {
			for (const set of [true, false]) {
				for (const use of ['Main = x + 1', 'Main = x & "a"']) {
					expect(found(type, set, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^collection-operand: .*Type mismatch/)]);
				}
				expect(found(type, set, 'x = "b"'), type).toEqual([expect.stringMatching(/^invalid-property-use: .*Invalid use of property/)]);
			}
		}
	});

	it('raises what DAO raises read whole', () => {
		expect(found('Recordset', true, 'Main = x')).toEqual([expect.stringMatching(/^object-default-value: .*'3001'/)]);
		expect(found('Database', true, 'Main = x')).toEqual([expect.stringMatching(/^object-default-value: .*'3001'/)]);
		expect(found('TableDef', true, 'Main = x')).toEqual([expect.stringMatching(/^object-default-value: .*'450'/)]);
	});
});

describe('Fields and TableDefs, whose default member Item needs an index', () => {
	it('are Argument not optional as an operand, set or not', () => {
		for (const type of ['Fields', 'TableDefs']) {
			for (const set of [true, false]) {
				for (const use of ['Main = x + 1', 'Main = x & "a"']) {
					expect(found(type, set, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^collection-operand: .*Argument not optional/)]);
				}
				expect(found(type, set, 'x = "b"'), type).toEqual([expect.stringMatching(/^set-required: .*Invalid use of property/)]);
			}
		}
	});

	it('raise 3001 read whole', () => {
		expect(found('Fields', true, 'Main = x')).toEqual([expect.stringMatching(/^object-default-value: .*'3001'/)]);
		expect(found('TableDefs', true, 'Main = x')).toEqual([expect.stringMatching(/^object-default-value: .*'3001'/)]);
	});
});

describe('a DAO Field', () => {
	it('raises 91 never set', () => {
		for (const use of ['Main = x + 1', 'Main = x & "a"', 'x = "b"', 'Main = x']) {
			expect(found('Field', false, use), use).toEqual([expect.stringMatching(/^object-variable-not-set: .*'91'/)]);
		}
	});

	it('gives its Value once set', () => {
		expect(found('Field', true, 'Main = x + 1')).toEqual([]);
		expect(found('Field', true, 'Main = x & "a"')).toEqual([]);
	});

	it('leaves members read by name alone', () => {
		expect(found('Recordset', true, 'Main = x.Fields(0).Value')).toEqual([]);
		expect(found('Recordset', true, 'Main = x(0)')).toEqual([]);
		expect(found('Recordset', true, 'Main = x!a')).toEqual([]);
		expect(found('TableDefs', true, 'Main = x.Count')).toEqual([]);
	});
});
