import { describe, it, expect } from 'vitest';
import {
	analyzeModule,
	VbaDiagnostic,
	DIAGNOSTIC_RULES,
	ProjectIndex,
	type HostObjectModel,
} from '../src/analyzer';

/** Returns the diagnostics whose code matches `code`. */
function byCode(diags: VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((d) => d.code === code);
}

/** Resolves the source substring a diagnostic span covers. */
function spanText(source: string, d: VbaDiagnostic): string {
	return source.slice(d.span.start, d.span.end);
}

function projectProcedures(
	modules: Array<{ moduleName: string; source: string }>,
): ReturnType<ProjectIndex['procedureSignatures']> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: 'standard',
			source: mod.source,
		});
	}
	return project.procedureSignatures();
}

function projectClassMembers(
	modules: Array<{
		moduleName: string;
		source: string;
		moduleKind?: 'standard' | 'class' | 'document' | 'userform';
	}>,
): ReturnType<ProjectIndex['projectClassMembers']> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: mod.moduleKind ?? 'standard',
			source: mod.source,
		});
	}
	return project.projectClassMembers();
}

function visibleProjectProcedures(
	modules: Array<{
		moduleName: string;
		source: string;
		moduleKind?: 'standard' | 'class' | 'document' | 'userform';
	}>,
	currentModule: string,
): ReadonlySet<string> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: mod.moduleKind ?? 'standard',
			source: mod.source,
		});
	}
	return project.visibleProcedureNames(currentModule);
}

function visibleProjectIdentifiers(
	modules: Array<{
		moduleName: string;
		source: string;
		moduleKind?: 'standard' | 'class' | 'document' | 'userform';
	}>,
	currentModule: string,
): ReadonlySet<string> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: mod.moduleKind ?? 'standard',
			source: mod.source,
		});
	}
	return project.visibleIdentifierNames(currentModule);
}

function visibleProjectTypes(
	modules: Array<{
		moduleName: string;
		source: string;
		moduleKind?: 'standard' | 'class' | 'document' | 'userform';
	}>,
	currentModule: string,
): ReturnType<ProjectIndex['visibleTypeNames']> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: mod.moduleKind ?? 'standard',
			source: mod.source,
		});
	}
	return project.visibleTypeNames(currentModule);
}

function visibleProjectNonTypeNames(
	modules: Array<{
		moduleName: string;
		source: string;
		moduleKind?: 'standard' | 'class' | 'document' | 'userform';
	}>,
	currentModule: string,
): ReadonlySet<string> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: mod.moduleKind ?? 'standard',
			source: mod.source,
		});
	}
	return project.visibleNonTypeNames(currentModule);
}

describe('analyzeModule - unterminated string', () => {
	it('flags a string with no closing quote', () => {
		const src = 'Sub T()\n    MsgBox "hello\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unterminated-string');
		expect(hits.length).toBe(1);
		expect(spanText(src, hits[0])).toBe('"hello');
		expect(hits[0].severity).toBe('error');
	});

	it('accepts a properly closed string, including doubled-quote escapes', () => {
		const src = 'Sub T()\n    MsgBox "say ""hi"" now"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(0);
	});

	it('treats a trailing escaped pair without a real close as unterminated', () => {
		const src = 'Sub T()\n    x = "ab""\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(1);
	});
});

describe('analyzeModule - duplicate procedures', () => {
	it('flags two Subs with the same name', () => {
		const src = 'Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-procedure');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Foo');
	});

	it('flags a Function colliding with a Sub of the same name', () => {
		const src = 'Sub Foo()\nEnd Sub\nFunction Foo()\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(1);
	});

	it('allows Property Get/Let/Set to share a name', () => {
		const src =
			'Property Get Item() As Long\nEnd Property\n' +
			'Property Let Item(v As Long)\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
	});

	it('flags a duplicate Property Get', () => {
		const src =
			'Property Get Item() As Long\nEnd Property\n' +
			'Property Get Item() As Long\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(1);
	});
});

describe('analyzeModule - duplicate declarations in scope', () => {
	it('flags the same local declared twice', () => {
		const src = 'Sub T()\n    Dim x As Long\n    Dim x As String\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-declaration');
		expect(hits).toHaveLength(1);
	});

	it('flags a local colliding with a parameter', () => {
		const src = 'Sub T(ByVal n As Long)\n    Dim n As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('flags duplicate parameter names', () => {
		const src = 'Sub T(a As Long, a As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('treats procedure scope as flat across branches', () => {
		const src =
			'Sub T()\n' +
			'    If True Then\n        Dim x As Long\n    End If\n' +
			'    Dim x As Long\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('does not flag distinct local names', () => {
		const src = 'Sub T()\n    Dim x As Long\n    Dim y As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate module members', () => {
	it('flags the same module variable declared twice', () => {
		const src = 'Private Total As Long\nPublic Total As String\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(1);
	});

	it('does not flag distinct module variables', () => {
		const src = 'Private A As Long\nPrivate B As Long\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});
});

describe('analyzeModule - assignment to constant', () => {
	it('flags assigning to a module-level Const', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    MAX = 5\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'const-assignment');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MAX');
	});

	it('flags assigning to a local Const', () => {
		const src = 'Sub T()\n    Const PI As Double = 3.14\n    PI = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(1);
	});

	it('does not flag comparing a Const in a condition', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    If MAX = 10 Then Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('does not flag assigning to a non-constant variable', () => {
		const src = 'Sub T()\n    Dim x As Long\n    x = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('does not flag a member or indexed left-hand side that shares a Const name', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    obj.MAX = 5\n    arr(MAX) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});
});

describe('analyzeModule - Option Explicit', () => {
	it('warns when a code module omits Option Explicit', () => {
		const src = 'Sub T()\n    Dim x As Long\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'option-explicit-missing');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('warning');
	});

	it('is silent when Option Explicit is present', () => {
		const src = 'Option Explicit\n\nSub T()\n    Dim x As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'option-explicit-missing')).toHaveLength(0);
	});

	it('is silent for an attribute-only / empty module', () => {
		const src = 'Attribute VB_Name = "Sheet1"\n';
		expect(byCode(analyzeModule(src), 'option-explicit-missing')).toHaveLength(0);
	});

	it('respects a severity override', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severities: { optionExplicitMissing: 'error' } }),
			'option-explicit-missing',
		);
		expect(hits[0].severity).toBe('error');
	});

	it('flags an undeclared bare assignment target when Option Explicit is present', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    notDeclared = ThisWorkbook.CanCheckIn()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('notDeclared');
		expect(hits[0].severity).toBe('error');
	});

	it('allows implicit Variant assignment when Option Explicit is absent', () => {
		const src = 'Sub T()\n    notDeclared = ThisWorkbook.CanCheckIn()\nEnd Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts declared local, module, parameter, and Function-return assignment targets', () => {
		const src =
			'Option Explicit\n' +
			'Private moduleValue As Long\n' +
			'Function T(ByVal arg As Long) As Long\n' +
			'    Dim localValue As Long\n' +
			'    localValue = 1\n' +
			'    moduleValue = localValue\n' +
			'    arg = moduleValue\n' +
			'    T = arg\n' +
			'End Function\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('flags an undeclared Set assignment target under Option Explicit', () => {
		const src = 'Option Explicit\nSub T()\n    Set notDeclared = ActiveSheet\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('notDeclared');
	});

	it('does not treat member or indexed assignment targets as bare undeclared variables yet', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Range("A1").Value = 1\n' +
			'    arr(1) = 2\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts exported standard-module globals visible through the project index', () => {
		const src = 'Option Explicit\nSub T()\n    sharedValue = 1\nEnd Sub\n';
		const knownIdentifiers = visibleProjectIdentifiers(
			[
				{ moduleName: 'Caller', source: src },
				{ moduleName: 'Globals', source: 'Public sharedValue As Long\n' },
			],
			'Caller',
		);
		expect(byCode(analyzeModule(src, { knownIdentifiers }), 'undeclared-variable')).toHaveLength(0);
	});

	it('can be switched off', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severities: { optionExplicitMissing: 'off' } }),
			'option-explicit-missing',
		);
		expect(hits).toHaveLength(0);
	});
});

describe('analyzeModule - general contract', () => {
	it('every diagnostic rule declares category, VBE equivalence, and evidence kind', () => {
		const categories = new Set([
			'syntax',
			'lexer',
			'parser',
			'realtime-recovery',
			'declaration',
			'semantic',
			'project-symbol',
			'module-kind',
			'excel-host',
			'style',
		]);
		const evidenceKinds = new Set([
			'compile-error',
			'deterministic-runtime-error',
			'runtime-risk',
			'style-policy',
		]);
		for (const [name, rule] of Object.entries(DIAGNOSTIC_RULES)) {
			expect(categories.has(rule.category), name).toBe(true);
			expect(typeof rule.vbeCompileEquivalent, name).toBe('boolean');
			expect(evidenceKinds.has(rule.diagnosticKind), name).toBe(true);
			if (rule.vbeCompileEquivalent) {
				expect(rule.diagnosticKind, name).toBe('compile-error');
			}
		}
	});

	it('never throws on malformed input', () => {
		const samples = ['', 'Sub', 'End Sub', 'Dim', '"', 'If Then', ':::'];
		for (const s of samples) {
			expect(() => analyzeModule(s)).not.toThrow();
		}
	});

	it('every emitted code is in the rule catalogue', () => {
		const known = new Set(
			Object.values(DIAGNOSTIC_RULES).map((r) => r.code),
		);
		const src =
			'Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n' +
			'Const C As Long = 1\nSub Bar()\n    C = 2\n    MsgBox "x\nEnd Sub\n';
		for (const d of analyzeModule(src)) {
			expect(known.has(d.code)).toBe(true);
		}
	});

	it('produces a clean module with no diagnostics', () => {
		const src =
			'Option Explicit\n\n' +
			'Const MAX As Long = 10\n\n' +
			'Sub Greet(ByVal name As String)\n' +
			'    Dim msg As String\n' +
			'    msg = "Hello " & name\n' +
			'    MsgBox msg\n' +
			'End Sub\n';
		expect(analyzeModule(src)).toHaveLength(0);
	});
});

describe('analyzeModule - unknown call statement', () => {
	const opts = { knownProcedures: new Set<string>() };

	it('flags a bare identifier that resolves to nothing', () => {
		const src = 'Sub Main()\n    asdfjalsdkfjas\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdfjalsdkfjas');
		expect(hits[0].severity).toBe('error');
	});

	it('does not flag a call to a procedure in another module', () => {
		const src = 'Sub Main()\n    DoWork\nEnd Sub\n';
		const known = { knownProcedures: new Set(['dowork']) };
		expect(byCode(analyzeModule(src, known), 'unknown-call')).toHaveLength(0);
	});

	it('uses project visibility for calls to other standard modules', () => {
		const caller = 'Sub Main()\n    DoWork\n    Secret\nEnd Sub\n';
		const helpers =
			'Public Sub DoWork()\nEnd Sub\n' +
			'Private Sub Secret()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				knownProcedures: visibleProjectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Helpers', source: helpers },
				], 'Caller'),
			}),
			'unknown-call',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Secret');
	});

	it('does not treat class-module methods as bare cross-module procedures', () => {
		const caller = 'Sub Main()\n    Save\nEnd Sub\n';
		const customer = 'Public Sub Save()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				knownProcedures: visibleProjectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Customer', moduleKind: 'class', source: customer },
				], 'Caller'),
			}),
			'unknown-call',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Save');
	});

	it('does not flag a call to a Sub defined in the same module', () => {
		const src = 'Sub Main()\n    Helper\nEnd Sub\nSub Helper()\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag VBA runtime functions/statements used bare', () => {
		const src = 'Sub Main()\n    DoEvents\n    Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a host global / Application member used bare', () => {
		const src = 'Sub Main()\n    Calculate\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag an in-scope local variable used bare', () => {
		const src = 'Sub Main()\n    Dim total As Long\n    total\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not run at all when knownProcedures is omitted', () => {
		const src = 'Sub Main()\n    asdfjalsdkfjas\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unknown-call')).toHaveLength(0);
	});

	it('ignores assignments, member calls, and known arg-bearing calls', () => {
		const src =
			'Sub Main()\n' +
			'    x = 1\n' +
			'    Debug.Print x\n' +
			'    MsgBox "hi"\n' +
			'    Foo 1, 2\n' +
			'End Sub\n';
		const known = { knownProcedures: new Set(['foo']) };
		expect(byCode(analyzeModule(src, known), 'unknown-call')).toHaveLength(0);
	});

	it('does not report unknown-call for a dangling member-access dot', () => {
		const src = 'Sub Main()\n    aadf.\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('ignores a line label', () => {
		const src = 'Sub Main()\n    GoTo done\ndone:\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('flags an unknown parenless call with arguments', () => {
		const src = 'Sub Main()\n    msrbox ""\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('msrbox');
	});

	it('flags an unknown call with multiple arguments', () => {
		const src = 'Sub Main()\n    Frobnicate 1, 2, 3\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Frobnicate');
	});

	it('flags an unknown explicit Call statement', () => {
		const src = 'Sub Main()\n    Call DoesNotExist(1)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
	});

	it('does not flag a known procedure called with arguments', () => {
		const src = 'Sub Main()\n    DoWork 1, 2\nEnd Sub\n';
		const known = { knownProcedures: new Set(['dowork']) };
		expect(byCode(analyzeModule(src, known), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a runtime function called with arguments', () => {
		const src = 'Sub Main()\n    Debug.Print Left("abc", 1)\n    MsgBox "hi", vbOKOnly\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a bare host member used with parentheses', () => {
		const src =
			'Sub Main()\n' +
			'    Cells(1, 1).Select\n' +
			'    Range("A1").Value = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});
});

describe('analyzeModule - non-callable call statements', () => {
	it('flags a bare local variable statement rejected by VBE Compile', () => {
		const src = 'Sub Main()\n    Dim testStr As String\n    testStr\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('testStr');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('local variable');
	});

	it('flags a local variable used with call arguments', () => {
		const src = 'Sub Main()\n    Dim testStr As String\n    testStr "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('testStr');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('local variable');
	});

	it('flags an explicit Call to a parameter', () => {
		const src = 'Sub Main(ByVal testStr As String)\n    Call testStr\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('testStr');
		expect(hits[0].message).toContain('parameter');
	});

	it('does not flag callable procedures or runtime statements', () => {
		const src = 'Sub Main()\n    Helper "ok"\n    Beep\nEnd Sub\nSub Helper(ByVal s As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'non-callable-call')).toHaveLength(0);
	});
});

describe('analyzeModule - scalar member access', () => {
	it('flags a trailing dot on a local String variable', () => {
		const src = 'Sub Main()\n    Dim value As String\n    value.\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('String');
		expect(hits[0].message).toContain('Syntax error');
	});

	it('flags named member access on a local Integer variable', () => {
		const src = 'Sub Main()\n    Dim value As Integer\n    value.Length\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
		expect(hits[0].message).toContain('Integer');
		expect(hits[0].message).toContain('Invalid qualifier');
	});

	it('flags scalar parameters and module variables', () => {
		const src =
			'Private moduleName As String\n' +
			'Sub Main(ByVal count As Long)\n' +
			'    moduleName.Length\n' +
			'    count.Value\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['moduleName.', 'count.']);
	});

	it('treats colon-separated declaration and member access statements independently', () => {
		const src = 'Sub Main()\n    Dim value As String: value.Length\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
	});

	it('does not flag Variant, Object, or unknown receiver types', () => {
		const src =
			'Sub Main()\n' +
			'    Dim flexible As Variant\n' +
			'    Dim item As Object\n' +
			'    Dim person As Person\n' +
			'    flexible.\n' +
			'    item.\n' +
			'    person.\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid procedure header', () => {
	it('flags a procedure name that contains a space', () => {
		const src = 'Sub My Sub\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-proc-header');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Sub');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a junk token after a Function name', () => {
		const src = 'Function Calc Total() As Long\nEnd Function\n';
		const hits = byCode(analyzeModule(src), 'invalid-proc-header');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Total');
	});

	it('does not flag a valid parameterless Sub', () => {
		const src = 'Sub Run\n    Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a valid Sub with parameters', () => {
		const src = 'Sub Greet(ByVal who As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a Function with a return type', () => {
		const src = 'Function Total() As Long\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a Property Get with a return type', () => {
		const src = 'Property Get Name() As String\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});
});

describe('analyzeModule - reserved declaration names', () => {
	it('flags reserved keywords used as procedure and variable names', () => {
		const src =
			'Function Dim() As String\n' +
			'    Dim In As String\n' +
			'    Dim = "ok"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'invalid-declaration-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Dim', 'In']);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('flags reserved keywords used as type, enum, field, member, and parameter names', () => {
		const src =
			'Public Type Type\n' +
			'    For As String\n' +
			'End Type\n' +
			'Public Enum Enum\n' +
			'    In\n' +
			'End Enum\n' +
			'Public Sub T(ByVal New As String)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-declaration-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Type',
			'For',
			'Enum',
			'In',
			'New',
		]);
	});

	it('allows bracketed reserved words as foreign names', () => {
		const src =
			'Function [Dim]() As String\n' +
			'    Dim [In] As String\n' +
			'    [Dim] = [In]\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'invalid-declaration-name')).toHaveLength(0);
	});
});

describe('analyzeModule - unbalanced parentheses', () => {
	it('flags a missing closing parenthesis', () => {
		const src = 'Sub T()\n    x = (1 + 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('(');
		expect(hits[0].severity).toBe('error');
	});

	it('flags an unexpected closing parenthesis', () => {
		const src = 'Sub T()\n    x = 1 + 2)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe(')');
	});

	it('reports at most one diagnostic per statement', () => {
		const src = 'Sub T()\n    x = ((1 + 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(1);
	});

	it('does not flag balanced and nested parentheses', () => {
		const src =
			'Sub T()\n' +
			'    x = (1 + (2 * 3))\n' +
			'    Debug.Print Left("ab", 1)\n' +
			'    Cells(1, 1).Value = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('does not count parentheses inside strings or comments', () => {
		const src =
			'Sub T()\n' +
			'    x = "a (b c"  \' a ) in a comment\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('balances across a line continuation', () => {
		const src = 'Sub T()\n    x = (1 + _\n        2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('treats a colon-separated statement independently', () => {
		const src = 'Sub T()\n    a = (1 + 2) : b = 3)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe(')');
	});
});

describe('analyzeModule - argument count', () => {
	it('flags too few arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Greet');
		expect(hits[0].message).toContain('expected 2 arguments');
		expect(hits[0].message).toContain('got 1');
	});

	it('flags too many arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob", "Cat"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('got 3');
	});

	it('flags a lone-identifier call that omits required arguments', () => {
		const src =
			'Sub Main()\n' +
			'    Greet\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('got 0');
	});

	it('flags a runtime function call that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MsgBox');
		expect(hits[0].message).toContain('expected between 1 and 5 arguments');
		expect(hits[0].message).toContain('got 0');
	});

	it('flags a parenless runtime function statement that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MsgBox');
		expect(hits[0].message).toContain('got 0');
	});

	it('validates an explicit Call statement', () => {
		const src =
			'Sub Main()\n' +
			'    Call Greet("Ann", "Bob", "Cat")\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(1);
	});

	it('flags an omitted required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(total, )\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('TaxRate');
		expect(hits[0].message).toContain('Argument not optional');
	});

	it('flags an omitted leading required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(, 0.08)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('Subtotal');
	});

	it('accepts an omitted Optional argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, Optional ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(total, )\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('accepts a correct argument count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('honours Optional parameters', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'    Greet "Ann", "Bob"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, Optional ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('honours ParamArray (no upper bound)', () => {
		const src =
			'Sub Main()\n' +
			'    Log "a", "b", "c", "d"\n' +
			'End Sub\n' +
			'Sub Log(ParamArray items() As Variant)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('validates named-argument names but not the count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet who:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('who');
		expect(hits[0].message).toContain('Named argument not found');
	});

	it('accepts a valid named argument', () => {
		const src =
			'Sub Main()\n' +
			'    Greet a:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, Optional ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check calls to unknown or cross-module procedures', () => {
		const src =
			'Sub Main()\n' +
			'    SomethingElse 1, 2, 3\n' +
			'    MsgBox "hi", vbOKOnly, "title", 0, 0\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check runtime statements whose curated signature has no parameter list', () => {
		const src = 'Sub Main()\n    Randomize 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check an ambiguous (duplicated) procedure name', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('uses a unique exported project signature for cross-module argument count', () => {
		const caller =
			'Public Sub Main()\n' +
			'    PrintTotal 100\n' +
			'End Sub\n';
		const helper =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Helpers', source: helper },
				]),
			}),
			'argument-count',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('PrintTotal');
		expect(hits[0].message).toContain('expected 2 arguments');
	});

	it('does not arity-check ambiguous exported project signatures', () => {
		const caller =
			'Public Sub Main()\n' +
			'    PrintTotal 100\n' +
			'End Sub\n';
		const first = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const second =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures: projectProcedures([
						{ moduleName: 'Caller', source: caller },
						{ moduleName: 'First', source: first },
						{ moduleName: 'Second', source: second },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('uses a module-qualified project signature even when the bare name is ambiguous', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.PrintTotal 100\n' +
			'End Sub\n';
		const helpers =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const alternate = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Helpers', source: helpers },
					{ moduleName: 'Alternate', source: alternate },
				]),
			}),
			'argument-count',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('PrintTotal');
		expect(hits[0].message).toContain('expected 2 arguments');
	});

	it('keeps module-qualified project diagnostics stable under module order changes', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.PrintTotal 100\n' +
			'End Sub\n';
		const helpers =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const alternate = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const ordered = projectProcedures([
			{ moduleName: 'Caller', source: caller },
			{ moduleName: 'Helpers', source: helpers },
			{ moduleName: 'Alternate', source: alternate },
		]);
		const reversed = projectProcedures([
			{ moduleName: 'Alternate', source: alternate },
			{ moduleName: 'Helpers', source: helpers },
			{ moduleName: 'Caller', source: caller },
		]);
		const messages = (projectProcedures: ReturnType<ProjectIndex['procedureSignatures']>) =>
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures,
				}),
				'argument-count',
			).map((hit) => `${spanText(caller, hit)}:${hit.message}`);
		expect(messages(ordered)).toEqual(messages(reversed));
	});

	it('does not validate a module-qualified private project procedure', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.Hidden 100\n' +
			'End Sub\n';
		const helpers =
			'Private Sub Hidden(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures: projectProcedures([
						{ moduleName: 'Caller', source: caller },
						{ moduleName: 'Helpers', source: helpers },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('flags missing required arguments on generated host member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Dim wb As Workbook\n' +
			'    Set wb = Workbooks.Open()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Open');
		expect(hits[0].message).toContain('expected between 1 and 15 arguments');
		expect(hits[0].message).toContain('got 0');
	});

	it('flags extra arguments on generated host member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Call Application.Calculate(1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Calculate');
		expect(hits[0].message).toContain('expected 0 arguments');
		expect(hits[0].message).toContain('got 1');
	});

	it('flags missing required arguments on host members with fallback signatures', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Range\n' +
			'    Set r = ActiveSheet.Range()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Range');
		expect(hits[0].message).toContain('expected between 1 and 2 arguments');
		expect(hits[0].message).toContain('got 0');
	});

	it('does not treat collection indexing as member-call arity', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Range\n' +
			'    Set r = Workbooks(1).Sheets(1).Range()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Range');
		expect(hits[0].message).toContain('got 0');
	});

	it('accepts correct host member argument counts', () => {
		const src =
			'Sub Main()\n' +
			'    Dim wb As Workbook\n' +
			'    Application.Calculate\n' +
			'    Call Application.Calculate()\n' +
			'    Set wb = Workbooks.Open("Book1.xlsx")\n' +
			'    ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags missing required arguments on source-backed class member calls', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    Call p.Save()\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-count',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Save');
		expect(hits[0].message).toContain('expected 1 argument');
		expect(hits[0].message).toContain('got 0');
	});

	it('flags missing required arguments on current class Me member calls', () => {
		const src =
			'Public Sub Main()\n' +
			'    Call Me.Save()\n' +
			'End Sub\n' +
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Person',
				moduleKind: 'class',
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: src },
				]),
			}),
			'argument-count',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save');
		expect(hits[0].message).toContain('expected 1 argument');
	});
});

describe('analyzeModule - argument type validation', () => {
	it('flags a nonnumeric string literal passed to a same-module numeric parameter', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('Subtotal');
		expect(hits[0].message).toContain('Currency');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('uses a unique exported project signature for cross-module argument types', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Invoices', source: invoices },
				]),
			}),
			'argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('Subtotal');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('uses module-qualified project signatures for argument types', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = Invoices.InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Invoices', source: invoices },
				]),
			}),
			'argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('Subtotal');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('uses a unique exported project function return type in nested calls', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel()\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('MakeLabel');
		expect(hits[0].message).toContain('MakeLabel(...) As String');
	});

	it('uses a module-qualified project function return type in nested calls', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Labels.MakeLabel()\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('MakeLabel');
		expect(hits[0].message).toContain('Labels.MakeLabel(...) As String');
	});

	it('accepts numeric literals and numeric string literals for numeric parameters', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    a = InvoiceTotal(100, 0.08)\n' +
			'    b = InvoiceTotal("100", 0.08)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn on string variables whose runtime value is unknown', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim label As String\n' +
			'    total = InvoiceTotal(label)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn on Variant arguments whose runtime value is unknown', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim value As Variant\n' +
			'    total = InvoiceTotal(value)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('maps named arguments to the named parameter type', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal(TaxRate:=0.08, Subtotal:="blah")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('Subtotal');
	});

	it('validates selected native VBA runtime parameter types', () => {
		const src = 'Sub T()\n    x = Left("abcdef", "bad")\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('Length');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('does not infer runtime parameter types from display names', () => {
		const src = 'Sub T()\n    Randomize "bad"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn when a type is unknown or Variant-like', () => {
		const src = 'Sub T()\n    x = Format("blah", "0.00")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('uses same-module function return types as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('MakeLabel');
		expect(hits[0].message).toContain('MakeLabel(...) As String');
	});

	it('uses generated host member signatures for scalar argument types', () => {
		const src =
			'Sub T()\n' +
			'    Call Application.DeleteCustomList("bad")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('ListNum');
		expect(hits[0].message).toContain('Long');
	});

	it('uses source-backed class member signatures for argument types', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.NeedsObject("bad")\n' +
			'End Sub\n';
		const person =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('item');
		expect(hits[0].message).toContain('Object');
	});

	it('uses nested same-module function return types as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Function EchoLabel(ByVal value As String) As String\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject EchoLabel(MakeLabel())\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('EchoLabel');
		expect(hits[0].message).toContain('EchoLabel(...) As String');
	});

	it('uses curated runtime conversion function return types as argument types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject CStr(123)\n' +
			'    NeedsObject CDbl("1")\n' +
			'    NeedsObject CCur(1)\n' +
			'    NeedsObject CLng(1)\n' +
			'    NeedsObject CBool("True")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(5);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'CStr',
			'CDbl',
			'CCur',
			'CLng',
			'CBool',
		]);
	});

	it('uses obvious numeric arithmetic expression return types as argument types', () => {
		const src =
			'Public Function TaxRate() As Double\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim subtotal As Currency\n' +
			'    Dim fee As Double\n' +
			'    NeedsObject subtotal + fee * TaxRate()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('subtotal + fee * TaxRate()');
		expect(hits[0].message).toContain('numeric expression');
	});

	it('does not infer arithmetic expressions with unknown, Variant, or string operands', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Double\n' +
			'    Dim flexible As Variant\n' +
			'    Dim label As String\n' +
			'    NeedsObject amount + flexible\n' +
			'    NeedsObject amount + UnknownValue\n' +
			'    NeedsObject amount + label\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses string concatenation expression return types as argument types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim prefix As String\n' +
			'    Dim amount As Double\n' +
			'    NeedsObject prefix & amount & CStr(123)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('prefix & amount & CStr(123)');
		expect(hits[0].message).toContain('string concatenation expression');
	});

	it('does not infer string concatenation expressions with unknown or Variant operands', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim prefix As String\n' +
			'    Dim flexible As Variant\n' +
			'    NeedsObject prefix & flexible\n' +
			'    NeedsObject prefix & UnknownValue\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - assignment type validation', () => {
	it('errors on a nonnumeric string literal assigned to a numeric variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim total As Double\n' +
			'    total = "blah"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('total');
		expect(hits[0].message).toContain('Double');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('errors on a non-Boolean string literal assigned to a Boolean variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim enabled As Boolean\n' +
			'    enabled = "maybe"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"maybe"');
		expect(hits[0].message).toContain('Boolean');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('accepts VBA scalar coercions and unknown assignment values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim total As Double\n' +
			'    Dim label As String\n' +
			'    Dim flexible As Variant\n' +
			'    total = "100"\n' +
			'    label = 123\n' +
			'    total = flexible\n' +
			'    total = UnknownValue\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('requires Set when assigning to a known object variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-required');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ws');
		expect(hits[0].message).toContain('requires Set');
	});

	it('reports missing Set for Object variables without treating it as scalar coercion', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    item = "blah"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'set-required')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('checks scalar Function return assignment types', () => {
		const src =
			'Public Function Total() As Double\n' +
			'    Total = "blah"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('Total');
		expect(hits[0].message).toContain('Double');
	});

	it('requires Set for object Function return assignments', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    MakePerson = New Person\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MakePerson');
	});

	it('checks object Function return assignment compatibility', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Class1\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Class1');
		expect(hits[0].message).toContain('MakePerson');
	});

	it('accepts compatible object Function return assignment', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Person\n' +
			'End Function\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: '' },
			]),
		});
		expect(byCode(diagnostics, 'set-required')).toHaveLength(0);
		expect(byCode(diagnostics, 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('checks Property Get return assignments like Function returns', () => {
		const src =
			'Public Property Get Child() As Person\n' +
			'    Child = New Person\n' +
			'End Property\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
	});

	it('flags Set assignment between incompatible project class types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p = New Class1\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Class1');
		expect(hits[0].message).toContain('Person');
		expect(hits[0].message).toContain('New Class1');
	});

	it('accepts Set assignment to a project interface implemented by another class', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Class1\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: 'Public Sub Save()\nEnd Sub\n' },
				{ moduleName: 'Class1', moduleKind: 'class', source: 'Implements Person\n' },
			]),
		});
		expect(byCode(diagnostics, 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('flags scalar Set assignment to a project object variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = "bad"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('object value');
	});

	it('errors on a nonnumeric string literal assigned to a typed class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "blah"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('p.Age');
		expect(hits[0].message).toContain('Integer');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('accepts numeric string assignment to a typed class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "2"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'assignment-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(0);
	});

	it('checks assignment types for public class fields', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "blah"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('p.Age');
		expect(hits[0].message).toContain('Integer');
	});

	it('accepts compatible assignments to public class fields', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "2"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'assignment-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(0);
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('requires Set for object-valued public class fields', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
		expect(hits[0].message).toContain('p.Child');
	});

	it('accepts Set for object-valued public class fields', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Person\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'set-required')).toHaveLength(0);
		expect(byCode(diagnostics, 'set-requires-object')).toHaveLength(0);
	});

	it('requires Set for Property Set object members', () => {
		const person =
			'Private mChild As Person\n' +
			'Public Property Set Child(ByVal value As Person)\n' +
			'    Set mChild = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
	});

	it('flags Set used against scalar source-backed members', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-requires-object',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Age');
		expect(hits[0].message).toContain('Integer');
	});

	it('flags Set assignment between incompatible source-backed object member types', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Class1\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Class1');
		expect(hits[0].message).toContain('p.Child');
	});

	it('errors on assignment to a read-only class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'readonly-member-assignment',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Age');
		expect(hits[0].message).toContain('read-only');
	});

	it('errors when a known class receiver uses an unknown member in assignment', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Height = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Height');
		expect(hits[0].message).toContain('Person.Height');
	});

	it('errors when a known class receiver calls an unknown method', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Delete');
	});

	it('errors when current class Me uses an unknown member', () => {
		const src =
			'Public Sub Save()\n' +
			'    Me.Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Person',
				moduleKind: 'class',
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: src },
				]),
			}),
			'member-not-found',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Delete');
		expect(hits[0].message).toContain('Person.Delete');
	});

	it('accepts known class members and ambiguous project receiver types', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'    p.Unknown = 2\n' +
			'End Sub\n';
		const knownDiagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(knownDiagnostics, 'member-not-found')).toHaveLength(1);
		expect(spanText(src, byCode(knownDiagnostics, 'member-not-found')[0])).toBe(
			'Unknown',
		);
		const ambiguous = projectClassMembers([
			{ moduleName: 'Person', moduleKind: 'class', source: person },
			{ moduleName: 'Other', moduleKind: 'class', source: person.replace(/Age/g, 'Size') },
		]).map((type) => ({ ...type, name: 'Person' }));
		expect(
			byCode(analyzeModule(src, { projectClassMembers: ambiguous }), 'member-not-found'),
		).toHaveLength(0);
	});

	it('does not use source-only document module members to prove absence', () => {
		const workbook =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim wb As ThisWorkbook\n' +
			'    wb.DoesntExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: workbook },
			]),
		});
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('errors when ThisWorkbook uses a member absent from source and the exhaustive Workbook host surface', () => {
		const workbook =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.doesnotexist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: workbook },
			]),
		});
		const hits = byCode(diagnostics, 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('doesnotexist');
		expect(hits[0].message).toContain('ThisWorkbook.doesnotexist');
	});

	it('does not treat Workbook events as callable object members', () => {
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.AfterSave True\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('AfterSave');
		expect(hits[0].message).toContain('Excel.Workbook.AfterSave');
	});

	it('accepts ThisWorkbook members from source and the exhaustive Workbook host surface', () => {
		const workbook =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.Hello\n' +
			'    ThisWorkbook.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: workbook },
			]),
		});
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('uses the exhaustive Workbook host surface for ActiveWorkbook', () => {
		const src =
			'Public Sub T()\n' +
			'    ActiveWorkbook.doesnotexist\n' +
			'    ActiveWorkbook.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: '' },
			]),
		});
		const hits = byCode(diagnostics, 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('doesnotexist');
		expect(hits[0].message).toContain('Excel.Workbook.doesnotexist');
	});

	it('uses the exhaustive Workbook host surface for declared Workbook variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.doesnotexist\n' +
			'    wb.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: '' },
			]),
		});
		const hits = byCode(diagnostics, 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('doesnotexist');
		expect(hits[0].message).toContain('Excel.Workbook.doesnotexist');
	});

	it('uses the exhaustive Worksheet host surface for ActiveSheet', () => {
		const src =
			'Public Sub T()\n' +
			'    ActiveSheet.asdf\n' +
			'    ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdf');
		expect(hits[0].message).toContain('Excel.Worksheet.asdf');
	});

	it('uses the current workbook Me host surface for ThisWorkbook modules', () => {
		const src =
			'Public Sub T()\n' +
			'    Me.asdf\n' +
			'    Me.AcceptAllChanges\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
			'member-not-found',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdf');
		expect(hits[0].message).toContain('Excel.Workbook.asdf');
	});

	it('uses the exhaustive Worksheet host surface for declared Worksheet variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = ActiveSheet\n' +
			'    ws.asdf\n' +
			'    ws.Range("A1")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdf');
		expect(hits[0].message).toContain('Excel.Worksheet.asdf');
	});

	it('does not prove missing members from late-bound Object or Variant receivers', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    obj.asdf\n' +
			'    flexible.asdf\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set-assignment refinement for hard late-bound member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Set obj = ActiveSheet\n' +
			'    Set flexible = Workbooks(1).Worksheets(1)\n' +
			'    obj.asdf\n' +
			'    flexible.asdf\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set-assignment refinement for hard late-bound member-call diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Set obj = ActiveSheet\n' +
			'    Set flexible = Workbooks(1).Worksheets(1)\n' +
			'    obj.Range()\n' +
			'    flexible.Range()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('uses an exhaustive host object model to prove a missing member', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					exhaustive: true,
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src =
			'Public Sub T()\n' +
			'    Thing.Missing\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src, { hostModel: model }), 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Missing');
		expect(hits[0].message).toContain('Test.Thing.Missing');
	});

	it('does not use a curated non-exhaustive host object model to prove absence', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src =
			'Public Sub T()\n' +
			'    Thing.Missing\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { hostModel: model }), 'member-not-found'),
		).toHaveLength(0);
	});
});

describe('analyzeModule - missing Function return assignment', () => {
	it('warns when a Function never assigns its return variable', () => {
		const src =
			'Public Function myFunction() As String\n' +
			'\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('warning');
		expect(spanText(src, hits[0])).toBe('myFunction');
		expect(hits[0].message).toContain('default value As String');
	});

	it('accepts scalar and object Function return assignments', () => {
		const src =
			'Public Function Label() As String\n' +
			'    Label = "ready"\n' +
			'End Function\n' +
			'\n' +
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Person\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('checks Property Get procedures and ignores Subs', () => {
		const src =
			'Public Property Get Name() As String\n' +
			'End Property\n' +
			'\n' +
			'Public Sub Refresh()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Name');
	});
});

describe('analyzeModule - string arithmetic coercion', () => {
	it('errors on a nonnumeric string literal inside a numeric assignment expression', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As Integer\n' +
			'    shouldErrorTest1 = 1 + "string"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'string-arithmetic-coercion');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('"string"');
		expect(hits[0].message).toContain('shouldErrorTest1');
		expect(hits[0].message).toContain('Integer');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('does not warn on numeric strings or unknown arithmetic operands', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim n As Integer\n' +
			'    n = 1 + "2"\n' +
			'    n = 1 + UnknownValue\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(0);
	});

	it('accepts plus between string literals assigned to a String variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As String\n' +
			'    shouldErrorTest1 = "string" + "string"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - As type name validation', () => {
	it('flags runtime functions used as declaration type names', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest3 As Int\n' +
			'    shouldErrorTest3 = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-as-type-name');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('Int');
	});

	it('uses project and host type resolution across type-name positions', () => {
		const src =
			'Public Type Payload\n' +
			'    Owner As Person\n' +
			'End Type\n' +
			'\n' +
			'Public Function Make(ByVal state As Status, ByVal sheet As Worksheet) As Person\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    If TypeOf p Is Person Then Debug.Print "ok"\n' +
			'    Set Make = p\n' +
			'End Function\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Types', source: 'Public Enum Status\n    Active\nEnd Enum\n' },
		];

		expect(byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
		}), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags runtime functions used in shared type-name positions', () => {
		const src =
			'Public Type Bag\n' +
			'    Field As Left\n' +
			'End Type\n' +
			'\n' +
			'Public Function Make(ByVal item As Int) As Right\n' +
			'    Dim value As Mid\n' +
			'    Set value = New Left\n' +
			'    If TypeOf value Is Right Then Debug.Print "bad"\n' +
			'End Function\n';

		const hits = byCode(analyzeModule(src), 'invalid-as-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Left',
			'Int',
			'Right',
			'Mid',
			'Left',
			'Right',
		]);
	});

	it('lets a project type shadow a runtime function name', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim value As Int\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Module1', source: src },
			{ moduleName: 'Int', moduleKind: 'class' as const, source: '' },
		];

		expect(byCode(analyzeModule(src, {
			moduleName: 'Module1',
			projectTypes: visibleProjectTypes(modules, 'Module1'),
		}), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags reserved identifiers used in shared type-name positions', () => {
		const src =
			'Implements Return\n' +
			'Public Sub T()\n' +
			'    Dim value As Dim\n' +
			'    Set value = New For\n' +
			'    If TypeOf value Is In Then Debug.Print "bad"\n' +
			'End Sub\n';

		const hits = byCode(analyzeModule(src), 'invalid-as-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Return',
			'Dim',
			'For',
			'In',
		]);
	});

	it('flags visible project declarations that are known not to be types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As DoWork\n' +
			'    Dim b As SharedValue\n' +
			'    Dim c As Active\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{
				moduleName: 'Helpers',
				source:
					'Public Sub DoWork()\nEnd Sub\n' +
					'Public SharedValue As Long\n' +
					'Public Enum Status\n    Active\nEnd Enum\n',
			},
		];

		const hits = byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'DoWork',
			'SharedValue',
			'Active',
		]);
	});

	it('lets a project type shadow a visible non-type declaration name', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim value As Person\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Helpers', source: 'Public Sub Person()\nEnd Sub\n' },
		];

		expect(byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags ambiguous visible project type names', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim state As Status\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Status', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Types', source: 'Public Enum Status\n    Active\nEnd Enum\n' },
		];

		const hits = byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Status');
		expect(hits[0].message).toContain('ambiguous');
	});

	it('flags duplicate visible project type names even when their kind matches', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Payload\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'TypesA', source: 'Public Type Payload\n    Id As Long\nEnd Type\n' },
			{ moduleName: 'TypesB', source: 'Public Type Payload\n    Name As String\nEnd Type\n' },
		];

		const hits = byCode(analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		}), 'invalid-as-type-name');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Payload');
		expect(hits[0].message).toContain('ambiguous');
	});

	it('defers broad unknown type names to the project-wide binder', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest2 As Intege\n' +
			'    shouldErrorTest2 = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-as-type-name')).toHaveLength(0);
	});
});

describe('analyzeModule - Set assignment validation', () => {
	it('flags Set assignment to a known scalar variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest5 As Integer\n' +
			'    Set shouldErrorTest5 = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-requires-object');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('shouldErrorTest5');
	});

	it('does not flag Set assignment to Object, Variant, or unknown object types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Dim sheet As Worksheet\n' +
			'    Set item = Nothing\n' +
			'    Set flexible = Nothing\n' +
			'    Set sheet = Nothing\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'set-requires-object')).toHaveLength(0);
	});
});

describe('analyzeModule - declaration initializer', () => {
	it('flags a module-level Dim with a VB.NET-style initializer', () => {
		const src = 'Private x As Long = 1\n';
		const hits = byCode(analyzeModule(src), 'dim-initializer');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('=');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a local Dim with an initializer', () => {
		const src = 'Sub T()\n    Dim n As Long = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(1);
	});

	it('flags an initializer without an As clause', () => {
		const src = 'Sub T()\n    Dim n = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(1);
	});

	it('does not flag a Const declaration', () => {
		const src = 'Const Pi As Double = 3.14\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});

	it('does not flag a plain declaration', () => {
		const src = 'Sub T()\n    Dim n As Long\n    n = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});

	it('does not flag an array bound that contains digits', () => {
		const src = 'Sub T()\n    Dim a(1 To 10) As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});
});

describe('analyzeModule - unexpected declaration tokens', () => {
	it('flags a bare identifier after a complete local As type', () => {
		const src = 'Sub T()\n    Dim s1 As String thisshoulderror\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unexpected-declaration-token');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('thisshoulderror');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('will fail to compile');
	});

	it('flags trailing tokens in module declarations, parameters, and Type fields', () => {
		const src =
			'Private moduleName As String junk\n' +
			'Private Type Customer\n' +
			'    Name As String extra\n' +
			'End Type\n' +
			'Sub T(ByVal label As String trailing)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unexpected-declaration-token');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'junk',
			'extra',
			'trailing',
		]);
	});

	it('accepts normal declaration separators and parameter defaults', () => {
		const src =
			'Private moduleName As String\n' +
			'Sub T(Optional ByVal label As String = "ok")\n' +
			'    Dim first As String, second As Long: Dim third As Variant\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('does not guess inside fixed-length string declarations or qualified type names', () => {
		const src =
			'Private fixedName As String * 10\n' +
			'Sub T()\n' +
			'    Dim workbook As Excel.Workbook\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});
});

describe('analyzeModule - object module public declaration restrictions', () => {
	it('flags public declarations that cannot be object-module members', () => {
		const src =
			'Public Const MaxRows As Long = 1000\n' +
			'Public Names() As String\n' +
			'Public FixedName As String * 20\n' +
			'Public Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
			'object-module-public-member',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'MaxRows',
			'Names',
			'FixedName',
			'Customer',
			'Sleep',
		]);
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('object modules');
	});

	it('does not apply object-module public-member restrictions in standard modules', () => {
		const src =
			'Public Const MaxRows As Long = 1000\n' +
			'Public Names() As String\n' +
			'Public FixedName As String * 20\n' +
			'Public Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		expect(byCode(analyzeModule(src), 'object-module-public-member')).toHaveLength(0);
	});

	it('does not flag private object-module declarations in this public-member rule', () => {
		const src =
			'Private Const MaxRows As Long = 1000\n' +
			'Private Names() As String\n' +
			'Private FixedName As String * 20\n' +
			'Private Type Customer\n' +
			'    Name As String\n' +
			'End Type\n' +
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
				'object-module-public-member',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - event handler module scope guidance', () => {
	it('guides when a workbook handler is declared in a standard module', () => {
		const src = 'Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Module1', moduleKind: 'standard' }),
			'event-handler-module-scope',
		);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('information');
		expect(spanText(src, hits[0])).toBe('Workbook_Open');
		expect(hits[0].message).toContain('not where Excel wires that event');
	});

	it('does not guide for workbook handlers in ThisWorkbook', () => {
		const src = 'Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
				'event-handler-module-scope',
			),
		).toHaveLength(0);
	});

	it('does not guide for worksheet handlers in worksheet document modules', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n';
		expect(
			byCode(
				analyzeModule(src, { moduleName: 'Sheet1', moduleKind: 'document' }),
				'event-handler-module-scope',
			),
		).toHaveLength(0);
	});

	it('guides when a worksheet handler is declared in ThisWorkbook', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
			'event-handler-module-scope',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Worksheet_Change');
		expect(hits[0].message).toContain('workbook document module');
	});

	it('uses proven chart document subtype before giving guidance', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Worksheet_Calculate()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'RevenueChart',
				moduleKind: 'document',
				documentType: 'chart',
			}),
			'event-handler-module-scope',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Worksheet_Calculate');
		expect(hits[0].message).toContain('chart document module');
	});
});

describe('analyzeModule - Call requires parentheses', () => {
	it('flags an unparenthesised Call argument list', () => {
		const src = 'Sub T()\n    Call MsgBox "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"hello"');
		expect(hits[0].severity).toBe('error');
	});

	it('flags an unparenthesised member Call', () => {
		const src = 'Sub T()\n    Call obj.Method 1, 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(1);
	});

	it('flags a standalone zero-argument member call that uses empty parentheses without Call', () => {
		const src = 'Sub T()\n    ThisWorkbook.CanCheckIn()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('CanCheckIn()');
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a standalone zero-argument host method call with empty parentheses before arity checks', () => {
		const src = 'Sub T()\n    Application.Calculate()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a standalone zero-argument class method call with empty parentheses', () => {
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    p.Save()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save()');
	});

	it('flags a same-module zero-argument Function call statement with empty parentheses', () => {
		const src =
			'Sub mySub()\n' +
			'    myFunction()\n' +
			'End Sub\n' +
			'\n' +
			'Function myFunction() As String\n' +
			'    myFunction = "hello world!"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('myFunction()');
	});

	it('flags an unqualified same-class method call statement with empty parentheses', () => {
		const src =
			'Public Sub SaveAll()\n' +
			'    Save()\n' +
			'End Sub\n' +
			'\n' +
			'Public Sub Save()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
			'call-statement-forbids-parens',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save()');
	});

	it('flags an exported cross-module zero-argument Function call statement with empty parentheses', () => {
		const src =
			'Sub mySub()\n' +
			'    myFunction()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src, {
			moduleName: 'Caller',
			projectProcedures: projectProcedures([
				{ moduleName: 'Helpers', source: 'Public Function myFunction() As String\nEnd Function\n' },
			]),
		}), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('myFunction()');
	});

	it('validates non-empty standalone member call parentheses with known signatures', () => {
		const src = 'Sub T()\n    Application.Calculate(1)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(1);
	});

	it('accepts Call with parentheses', () => {
		const src = 'Sub T()\n    Call MsgBox("hello")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parameterless Call', () => {
		const src = 'Sub T()\n    Call DoWork\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts a Call to a parenthesised member chain', () => {
		const src = 'Sub T()\n    Call obj.Method(1, 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts parenthesized member calls with Call or in expression context', () => {
		const src =
			'Sub T()\n' +
			'    Dim ok As Boolean\n' +
			'    Call ThisWorkbook.CanCheckIn()\n' +
			'    ok = ThisWorkbook.CanCheckIn()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('accepts known Function calls with Call or in expression context', () => {
		const src =
			'Sub mySub()\n' +
			'    Dim value As String\n' +
			'    Call myFunction()\n' +
			'    value = myFunction()\n' +
			'End Sub\n' +
			'\n' +
			'Function myFunction() As String\n' +
			'    myFunction = "hello world!"\n' +
			'End Function\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not flag unknown bare empty-parentheses statements as project procedure calls', () => {
		const src = 'Sub T()\n    MaybeExternal()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not flag a parenless non-Call statement', () => {
		const src = 'Sub T()\n    MsgBox "hello"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - expression call requires parentheses', () => {
	it('flags a same-module Function called with parenless arguments in an assignment', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'expression-call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('InvoiceTotal');
	});

	it('flags a runtime Function called with parenless arguments in an assignment', () => {
		const src = 'Sub T()\n    answer = MsgBox "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'expression-call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MsgBox');
	});

	it('accepts a parenthesized Function call in an assignment', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal(100, 0.08)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parameterless Function reference in an expression', () => {
		const src =
			'Public Function CurrentTotal() As Currency\nEnd Function\n' +
			'Public Sub TestTotal()\n' +
			'    total = CurrentTotal + 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - parameter order', () => {
	it('flags a required parameter after an Optional one', () => {
		const src =
			'Sub T(Optional ByVal a As Long = 1, ByVal b As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'required-param-after-optional');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toContain('b');
	});

	it('accepts trailing Optional parameters', () => {
		const src =
			'Sub T(ByVal a As Long, Optional ByVal b As Long = 1)\nEnd Sub\n';
		expect(
			byCode(analyzeModule(src), 'required-param-after-optional'),
		).toHaveLength(0);
	});

	it('flags a ParamArray that is not last', () => {
		const src = 'Sub T(ParamArray items() As Variant, ByVal n As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-not-last');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toContain('items');
	});

	it('accepts a trailing ParamArray', () => {
		const src = 'Sub T(ByVal n As Long, ParamArray items() As Variant)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'paramarray-not-last')).toHaveLength(0);
	});
});

describe('analyzeModule - Exit statement matches procedure', () => {
	it('flags Exit Function inside a Sub', () => {
		const src = 'Sub T()\n    Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
		expect(hits[0].severity).toBe('error');
	});

	it('flags Exit Sub inside a Function', () => {
		const src = 'Function F() As Long\n    Exit Sub\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('flags Exit Sub inside a Property', () => {
		const src = 'Property Get Name() As String\n    Exit Sub\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('accepts a matching Exit Sub', () => {
		const src = 'Sub T()\n    Exit Sub\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('accepts a matching Exit Property', () => {
		const src = 'Property Let Name(v As String)\n    Exit Property\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('ignores Exit For and Exit Do', () => {
		const src =
			'Sub T()\n' +
			'    Do\n        Exit Do\n    Loop\n' +
			'    For i = 1 To 3\n        Exit For\n    Next i\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});
});

describe('analyzeModule - statement context', () => {
	it('flags an If statement missing Then', () => {
		const src = 'Sub T()\n    If x > 0\n        x = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'if-missing-then');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('If');
	});

	it('accepts multiline and single-line If statements with Then', () => {
		const src =
			'Sub T()\n' +
			'    If x > 0 Then\n        x = 1\n    End If\n' +
			'    If x > 1 Then x = 2\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'if-missing-then')).toHaveLength(0);
	});

	it('flags Case outside Select Case', () => {
		const src = 'Sub T()\n    Case 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'case-outside-select');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Case');
	});

	it('accepts Case inside Select Case', () => {
		const src =
			'Sub T()\n' +
			'    Select Case x\n' +
			'        Case 1, 2\n' +
			'            x = 3\n' +
			'        Case Else\n' +
			'            x = 4\n' +
			'    End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
	});

	it('flags leading-dot member access outside With', () => {
		const src = 'Sub T()\n    .Value = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'member-access-outside-with');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('accepts leading-dot member access inside With', () => {
		const src =
			'Sub T()\n' +
			'    With Range("A1")\n' +
			'        .Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
	});

	it('flags Exit For and Exit Do outside matching loops', () => {
		const src = 'Sub T()\n    Exit For\n    Exit Do\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-outside-block');
		expect(hits).toHaveLength(2);
		expect(spanText(src, hits[0])).toBe('Exit For');
		expect(spanText(src, hits[1])).toBe('Exit Do');
	});

	it('accepts Exit For and Exit Do inside matching loops', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n        Exit For\n    Next i\n' +
			'    Do\n        Exit Do\n    Loop\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-outside-block')).toHaveLength(0);
	});
});

describe('analyzeModule - Option placement', () => {
	it('flags an Option after a declaration', () => {
		const src = 'Private m_Count As Long\nOption Explicit\n';
		const hits = byCode(analyzeModule(src), 'option-after-declaration');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('flags an Option after a procedure', () => {
		const src = 'Sub T()\nEnd Sub\nOption Base 1\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(1);
	});

	it('accepts Options at the top of the module', () => {
		const src = 'Option Explicit\nOption Base 1\nPrivate m_Count As Long\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(0);
	});

	it('accepts Options after Attribute lines', () => {
		const src =
			'Attribute VB_Name = "Module1"\nOption Explicit\nSub T()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(0);
	});
});

