// Compiler-directive forms the VBE refuses (issue #130). Measured in Excel
// 16.0 (build 20326, 2026-09-26).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

describe('duplicate-const-directive and directive-trailing-statement (issue #130)', () => {
	it('flags a #Const defined twice in one module', () => {
		const src = 'Option Explicit\n#Const FEATURE = 1\n#Const FEATURE = 2\nFunction Main() As Variant\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), 'duplicate-const-directive', { span: 'FEATURE', message: 'Duplicate definition' });
	});

	it('flags code after the colon on a directive line', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n#If VBA7 Then: Debug.Print 1\n#End If\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), 'directive-trailing-statement', { span: 'Debug.Print 1', message: 'whole line' });
	});

	it('stays quiet for one #Const, a directive with a trailing comment and ordinary #If arms', () => {
		const src = "Option Explicit\n#Const FEATURE = 1\nFunction Main() As Variant\n#If FEATURE Then ' the feature\n    Main = 1\n#Else\n    Main = 2\n#End If\nEnd Function\n";
		expect(byCode(analyzeModule(src), 'duplicate-const-directive')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'directive-trailing-statement')).toHaveLength(0);
	});
});
