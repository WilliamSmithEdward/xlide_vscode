import { describe, it, expect } from 'vitest';
import {
	analyzeModule,
	VbaDiagnostic,
	DIAGNOSTIC_RULES,
	STRUCTURAL_DIAGNOSTIC_RULES,
	diagnosticMetadataForCode,
	diagnosticSourceForCode,
	isXlideDiagnosticSource,
	type HostObjectModel,
} from '../src/analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../src/vbaProjectAnalysis';

/** Returns the diagnostics whose code matches `code`. */
function byCode(diags: VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((d) => d.code === code);
}

/** Resolves the source substring a diagnostic span covers. */
function spanText(source: string, d: VbaDiagnostic): string {
	return source.slice(d.span.start, d.span.end);
}

type ProjectTestModule = VbaProjectModuleInput;

function projectOptions(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReturnType<typeof projectAnalysisOptionsForModule> {
	const project = buildVbaProjectIndex(modules);
	return projectAnalysisOptionsForModule(
		project,
		currentModule,
		projectProcedureSignatures(project),
	);
}

function analyzeProjectModule(
	source: string,
	modules: readonly ProjectTestModule[],
	currentModule: string,
	extra: Parameters<typeof analyzeModule>[1] = {},
): VbaDiagnostic[] {
	const current = modules.find(
		(mod) => mod.moduleName.toLowerCase() === currentModule.toLowerCase(),
	);
	const projectModules: readonly ProjectTestModule[] = [
		current
			? { ...current, moduleName: currentModule, source }
			: { moduleName: currentModule, source },
		...modules.filter(
			(mod) => mod.moduleName.toLowerCase() !== currentModule.toLowerCase(),
		),
	];
	return analyzeModule(source, {
		moduleName: currentModule,
		...projectOptions(projectModules, currentModule),
		...extra,
	});
}

function projectProcedures(
	modules: readonly ProjectTestModule[],
): NonNullable<ReturnType<typeof projectProcedureSignatures>> {
	const project = buildVbaProjectIndex(modules);
	return projectProcedureSignatures(project) ?? new Map();
}

function projectClassMembers(
	modules: readonly ProjectTestModule[],
): NonNullable<ReturnType<typeof projectAnalysisOptionsForModule>['projectClassMembers']> {
	const project = buildVbaProjectIndex(modules);
	return project.projectClassMembers();
}

function projectMemberSurfaces(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): NonNullable<ReturnType<typeof projectAnalysisOptionsForModule>['projectClassMembers']> {
	return projectOptions(modules, currentModule).projectClassMembers ?? [];
}

function visibleProjectProcedures(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReadonlySet<string> {
	return projectOptions(modules, currentModule).knownProcedures ?? new Set();
}

function visibleProjectIdentifiers(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReadonlySet<string> {
	return projectOptions(modules, currentModule).knownIdentifiers ?? new Set();
}

function visibleProjectTypes(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): NonNullable<ReturnType<typeof projectAnalysisOptionsForModule>['projectTypes']> {
	return projectOptions(modules, currentModule).projectTypes ?? [];
}

function visibleProjectNonTypeNames(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReadonlySet<string> {
	return projectOptions(modules, currentModule).knownNonTypeNames ?? new Set();
}

describe('vbaProjectAnalysis helper', () => {
	it('derives project member surfaces from the live current-module overlay', () => {
		const savedClass = [
			'Public Property Get Name() As String',
			'End Property',
		].join('\n');
		const liveClass = [
			'Public Property Get Age() As Long',
			'End Property',
		].join('\n');
		const project = buildVbaProjectIndex(
			[
				{ moduleName: 'Person', moduleKind: 'class', source: savedClass },
				{ moduleName: 'Caller', moduleKind: 'standard', source: '' },
			],
			{ moduleName: 'Person', moduleKind: 'class', source: liveClass },
		);
		const person = projectAnalysisOptionsForModule(project, 'Caller')
			.projectClassMembers
			?.find((surface) => surface.name === 'Person');

		expect(person?.members.map((member) => member.name)).toEqual(['Age']);
	});
});

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

describe('analyzeModule - invalid line continuation', () => {
	it('flags a continuation underscore followed by a trailing comment', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "hello" & _ \' bad trailing comment\n' +
			'        "world"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-line-continuation');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe("_ ' bad trailing comment");
		expect(hits[0].severity).toBe('error');
	});

	it('flags a continuation underscore followed by more code on the same line', () => {
		const src = 'Sub T()\n    total = 1 _ + 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-line-continuation');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('_ + 2');
	});

	it('flags a likely continuation with no whitespace before the underscore', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "hello" &_\n' +
			'        "world"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-line-continuation');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('_');
	});

	it('accepts valid continuations', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "hello" & _\n' +
			'        "world"\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'invalid-line-continuation')).toHaveLength(0);
	});

	it('ignores underscores in identifiers, strings, and comments', () => {
		const src =
			'Sub T()\n' +
			'    Dim value_name As Long\n' +
			'    value_name = 1\n' +
			'    Debug.Print "text _ inside string"\n' +
			"    ' comment _ at the end\n" +
			'    Rem comment _ at the end\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'invalid-line-continuation')).toHaveLength(0);
	});

	it('leaves a final dangling continuation as a realtime recovery case', () => {
		const src = 'Sub T()\n    Debug.Print "hello" & _';

		expect(byCode(analyzeModule(src), 'invalid-line-continuation')).toHaveLength(0);
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

	it('filters duplicate procedures in default VBA7 conditional branches', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#Else\n' +
			'Public Sub Configure()\nEnd Sub\n' +
			'#End If\n';

		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-procedure',
			),
		).toHaveLength(0);
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

	it('filters duplicate declarations in default VBA7 conditional branches', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim value As LongPtr\n' +
			'#Else\n' +
			'    Dim value As Long\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'duplicate-declaration',
			),
		).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-declaration',
			),
		).toHaveLength(0);
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

	it('filters duplicate module variables in default VBA7 conditional branches', () => {
		const src =
			'#If VBA7 Then\n' +
			'Private activeSheetPtr As LongPtr\n' +
			'#Else\n' +
			'Private activeSheetPtr As Long\n' +
			'#End If\n';

		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-module-variable',
			),
		).toHaveLength(0);
	});

	it('filters duplicate module variables in default Win64 conditional branches', () => {
		const src =
			'#If Win64 Then\n' +
			'Private nativeHandle As LongPtr\n' +
			'#Else\n' +
			'Private nativeHandle As Long\n' +
			'#End If\n';

		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { Win64: false } },
				}),
				'duplicate-module-variable',
			),
		).toHaveLength(0);
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

	it('flags assigning to a Const after a numeric line label', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n' +
			'10 MAX = 5\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'const-assignment');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MAX');
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
		const src = 'Option Explicit\nSub T()\n    MissingProc\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				knownProcedures: new Set<string>(),
				severityOverrides: { 'unknown-call': 'warning' },
			}),
			'unknown-call',
		);
		expect(hits[0].severity).toBe('warning');
	});

	it('ignores severity overrides that violate rule guardrails', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severityOverrides: { 'option-explicit-missing': 'error' } }),
			'option-explicit-missing',
		);
		expect(hits[0].severity).toBe('warning');
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

	it('accepts known member and indexed assignment receivers under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Range("A1").Value = 1\n' +
			'    declaredArr(1) = 2\n' +
			'End Sub\n';
		const knownIdentifiers = new Set<string>(['declaredarr']);
		expect(
			byCode(analyzeModule(src, { knownIdentifiers }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('flags undeclared identifiers read from assignment right-hand sides and call arguments', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim total As Long\n' +
			'    total = missingValue + 1\n' +
			'    MsgBox missingMessage\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['missingValue', 'missingMessage']);
	});

	it('flags undeclared identifiers in control-flow expressions and loop targets', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    If missingCondition Then Beep\n' +
			'    For i = 1 To maxCount\n' +
			'    Next i\n' +
			'    With missingObject\n' +
			'        .Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'missingCondition',
			'i',
			'maxCount',
			'missingObject',
		]);
	});

	it('flags undeclared member receivers and indexed bases under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    obj.Value = 1\n' +
			'    arr(ix) = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['obj', 'arr', 'ix']);
	});

	it('accepts declared reads, exported globals, and exported enum members', () => {
		const src =
			'Option Explicit\n' +
			'Sub T(ByVal arg As Long)\n' +
			'    Dim localValue As Long\n' +
			'    localValue = arg + sharedValue + SharedOnly\n' +
			'End Sub\n';
		const knownIdentifiers = visibleProjectIdentifiers(
			[
				{ moduleName: 'Caller', source: src },
				{
					moduleName: 'Globals',
					source:
						'Public sharedValue As Long\n' +
						'Public Enum SharedMode\n    SharedOnly\nEnd Enum\n',
				},
			],
			'Caller',
		);
		expect(byCode(analyzeModule(src, { knownIdentifiers }), 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts built-in VBA and Excel constants under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    MsgBox "hi", vbOKOnly\n' +
			'    Application.DisplayAlerts = vbFalse\n' +
			'    ActiveSheet.Range("A1").End(xlUp).Select\n' +
			'    Dim dashStyle As Long\n' +
			'    dashStyle = msoLineDash\n' +
			'    Err.Raise vbObjectError + 1, "Person.Age", "Age must be between 0 and 120"\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts the VBA namespace and compare aliases under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim seen As Scripting.Dictionary\n' +
			'    Set seen = New Scripting.Dictionary\n' +
			'    seen.CompareMode = TextCompare\n' +
			'    MsgBox VBA.CStr(1)\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts Erl in line-numbered error handlers under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'10 Dim message As String\n' +
			'20 On Error GoTo Handler\n' +
			'30 Err.Raise 5\n' +
			'40 Exit Sub\n' +
			'Handler:\n' +
			'50 message = "Error on line " & Erl\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('handles line-numbered assignment targets under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim declared As Long\n' +
			'    Dim obj As Object\n' +
			'10 declared = 1\n' +
			'20 Set obj = ActiveSheet\n' +
			'30 missing = 1\n' +
			'40 Set missingObj = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['missing', 'missingObj']);
		expect(hits.every((hit) => hit.message.includes('assigning to it'))).toBe(true);
	});

	it('does not flag type names, labels, named-argument labels, or unknown external-style calls as reads', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'done:\n' +
			'    Set p = New Person\n' +
			'    If TypeOf p Is Person Then GoTo done\n' +
			'    MsgBox Prompt:=p\n' +
			'    MaybeExternal missingArg\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>(['p']) }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts exported standard-module globals visible through the project index', () => {
		const src = 'Option Explicit\nSub T()\n    sharedValue = 1\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(
			src,
			[
				{ moduleName: 'Globals', source: 'Public sharedValue As Long\n' },
			],
			'Caller',
		);
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts module-qualified Function calls on a Set assignment right-hand side', () => {
		const caller =
			'Option Explicit\n' +
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    Set item = Factories.MakeItem()\n' +
			'    Set item = Factories.MakeOther\n' +
			'End Sub\n';
		const factories =
			'Public Function MakeItem() As Object\n' +
			'End Function\n' +
			'Public Function MakeOther() As Object\n' +
			'End Function\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Factories', source: factories },
		], 'Caller');
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts module-qualified standard-module values on expression right-hand sides', () => {
		const caller =
			'Option Explicit\n' +
			'Public Sub T()\n' +
			'    Dim value As Long\n' +
			'    value = Globals.SharedConst + Settings.SomePublicValue + Globals.SharedOnly\n' +
			'End Sub\n';
		const globals =
			'Public Const SharedConst As Long = 1\n' +
			'Public Enum SharedMode\n' +
			'    SharedOnly\n' +
			'End Enum\n';
		const settings = 'Public SomePublicValue As Long\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: globals },
			{ moduleName: 'Settings', source: settings },
		], 'Caller');
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('keeps unknown module-qualified value qualifiers visible under Option Explicit', () => {
		const caller =
			'Option Explicit\n' +
			'Public Sub T()\n' +
			'    Dim value As Long\n' +
			'    value = MissingModule.SharedConst\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [], 'Caller');
		const hits = byCode(diagnostics, 'undeclared-variable');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('MissingModule');
	});

	it('can be switched off', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severityOverrides: { 'option-explicit-missing': 'off' } }),
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
		for (const [name, rule] of Object.entries({
			...DIAGNOSTIC_RULES,
			...STRUCTURAL_DIAGNOSTIC_RULES,
		})) {
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

	it('resolves metadata for semantic and structural diagnostic codes', () => {
		expect(diagnosticMetadataForCode('undeclared-variable')).toMatchObject({
			category: 'project-symbol',
			vbeCompileEquivalent: true,
			diagnosticKind: 'compile-error',
		});
		expect(diagnosticMetadataForCode('missing-block-closer')).toMatchObject({
			title: STRUCTURAL_DIAGNOSTIC_RULES.missingBlockCloser.title,
			category: 'syntax',
			vbeCompileEquivalent: true,
			diagnosticKind: 'compile-error',
		});
		expect(diagnosticMetadataForCode('not-a-real-rule')).toBeUndefined();
	});

	it('uses metadata source labels for Problems filtering without changing rule codes', () => {
		expect(diagnosticSourceForCode('missing-block-closer')).toBe('XLIDE/VBE');
		expect(diagnosticSourceForCode('undeclared-variable')).toBe('XLIDE/VBE');
		expect(diagnosticSourceForCode('string-arithmetic-coercion')).toBe('XLIDE/runtime');
		expect(diagnosticSourceForCode('missing-return-assignment')).toBe('XLIDE/risk');
		expect(diagnosticSourceForCode('option-explicit-missing')).toBe('XLIDE/style');
		expect(diagnosticSourceForCode(undefined)).toBe('XLIDE');
		expect(isXlideDiagnosticSource('XLIDE')).toBe(true);
		expect(isXlideDiagnosticSource('XLIDE/VBE')).toBe(true);
		expect(isXlideDiagnosticSource('typescript')).toBe(false);
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
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helpers },
			], 'Caller'),
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
		expect(hits[0].data?.createProcedureStub).toMatchObject({
			procedureName: 'Frobnicate',
			edit: {
				span: { start: src.length, end: src.length },
				newText:
					'\nPrivate Sub Frobnicate(ByVal arg1 As Variant, ByVal arg2 As Variant, ByVal arg3 As Variant)\n' +
					'End Sub\n',
			},
		});
	});

	it('flags an unknown explicit Call statement', () => {
		const src = 'Sub Main()\n    Call DoesNotExist(1)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
		expect(hits[0].data?.createProcedureStub?.edit.newText).toContain(
			'Private Sub DoesNotExist(ByVal arg1 As Variant)\n',
		);
	});

	it('flags unknown calls after numeric line labels', () => {
		const src =
			'Sub Main()\n' +
			'10 DoesNotExist 1\n' +
			'20 Call AlsoMissing(2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DoesNotExist', 'AlsoMissing']);
	});

	it('does not attach procedure-stub metadata for omitted argument slots', () => {
		const src = 'Sub Main()\n    Call DoesNotExist(1, , 3)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
		expect(hits[0].data?.createProcedureStub).toBeUndefined();
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

	it('does not flag indexed object variables feeding member access', () => {
		const src =
			'Sub Main()\n' +
			'    Dim buckets As Object\n' +
			'    Call buckets("ready").Add(1)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'non-callable-call')).toHaveLength(0);
	});

	it('flags project-visible exported non-callables used as call statements', () => {
		const caller =
			'Sub Main()\n' +
			'    SharedValue\n' +
			'    MaxValue 1\n' +
			'    Call SharedMode\n' +
			'    Active\n' +
			'End Sub\n';
		const helpers =
			'Public SharedValue As Long\n' +
			'Public Const MaxValue As Long = 10\n' +
			'Public Enum SharedMode\n' +
			'    Active\n' +
			'End Enum\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', source: helpers },
		], 'Caller');
		const hits = byCode(diagnostics, 'non-callable-call');

		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedValue',
			'MaxValue',
			'SharedMode',
			'Active',
		]);
		expect(hits[0].message).toContain('module variable');
		expect(hits[1].message).toContain('constant');
		expect(hits[2].message).toContain('enum type');
		expect(hits[3].message).toContain('enum member');
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
	});

	it('keeps project non-callables silent when a visible procedure shares the name', () => {
		const caller = 'Sub Main()\n    SharedName\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedName As Long\n' },
			{ moduleName: 'Helpers', source: 'Public Sub SharedName()\nEnd Sub\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'non-callable-call')).toHaveLength(0);
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
	});

	it('keeps duplicate project non-callables silent instead of unknown', () => {
		const caller = 'Sub Main()\n    SharedName\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedName As Long\n' },
			{ moduleName: 'OtherGlobals', source: 'Public Const SharedName As Long = 1\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'non-callable-call')).toHaveLength(0);
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
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

describe('analyzeModule - invalid identifier starts', () => {
	it('flags digit-start variable and parameter names', () => {
		const src =
			'Sub T(ByVal 2bad As Long)\n' +
			'    Dim 1bad As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-identifier-start');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['2bad', '1bad']);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('flags digit-start module, procedure, and compiler-constant declaration names', () => {
		const src =
			'#Const 1debug = True\n' +
			'Public Declare Sub 9Sleep Lib "kernel32" ()\n' +
			'Public Type 3Thing\n' +
			'    4Field As Long\n' +
			'End Type\n' +
			'Public Enum 5Choice\n' +
			'    6First = 1\n' +
			'End Enum\n' +
			'Sub 7Run()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-identifier-start');

		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'1debug',
			'9Sleep',
			'3Thing',
			'4Field',
			'5Choice',
			'6First',
			'7Run',
		]);
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('accepts ordinary and bracketed foreign names', () => {
		const src =
			'Sub T(ByVal value1 As Long)\n' +
			'    Dim value2 As Long\n' +
			'    Dim [1bad] As Long\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'invalid-identifier-start')).toHaveLength(0);
	});
});

describe('analyzeModule - module declarations inside procedures', () => {
	it('flags module-only declarations inside procedure bodies', () => {
		const src =
			'Sub T()\n' +
			'    Option Explicit\n' +
			'    DefLng A-Z\n' +
			'    Public leakedValue As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'module-declaration-in-procedure');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Option', 'DefLng', 'Public']);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('checks nested procedure blocks and ignores inactive conditional branches', () => {
		const src =
			'Sub T()\n' +
			'    If True Then\n' +
			'        DefStr A-Z\n' +
			'    End If\n' +
			'    #If Enabled Then\n' +
			'        DefLng A-Z\n' +
			'    #End If\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { conditionalCompilation: { compilerConstants: { Enabled: false } } }),
			'module-declaration-in-procedure',
		);

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DefStr']);
	});

	it('accepts module declarations at module level and procedure-local declarations', () => {
		const src =
			'Option Explicit\n' +
			'DefLng A-Z\n' +
			'Private moduleValue As Long\n' +
			'Sub T()\n' +
			'    Static localValue As Long\n' +
			'    Const localConst As Long = 1\n' +
			'    Dim localDim As Long\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
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

	it('accepts Event declarations with parameter lists in object modules', () => {
		const src =
			'Public Event BeforeAdd(ByRef arr As Variant, ByRef cancel As Boolean)\n' +
			'Public Event AfterAdd(ByRef arr As Variant)\n';

		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			const diagnostics = analyzeModule(src, { moduleKind });
			expect(byCode(diagnostics, 'invalid-declaration-name')).toHaveLength(0);
			expect(byCode(diagnostics, 'unexpected-declaration-token')).toHaveLength(0);
			expect(byCode(diagnostics, 'duplicate-module-variable')).toHaveLength(0);
			expect(byCode(diagnostics, 'event-declaration-module-kind')).toHaveLength(0);
		}
	});

	it('accepts user-defined type fields named Type', () => {
		const src =
			'Private Type PICTDESC\n' +
			'    size As Long\n' +
			'    Type As Long\n' +
			'End Type\n';

		expect(byCode(analyzeModule(src), 'invalid-declaration-name')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
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
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'b',
			edit: {
				span: {
					start: src.indexOf('"Ann"') + '"Ann"'.length,
					end: src.indexOf('"Ann"') + '"Ann"'.length,
				},
				newText: ', TODO_b',
			},
		});
	});

	it('checks argument counts for calls after numeric line labels', () => {
		const src =
			'Sub Main()\n' +
			'10 Greet "Ann"\n' +
			'20 Call Greet("Ann")\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['Greet', 'Greet']);
		expect(hits.every((hit) => hit.message.includes('expected 2 arguments'))).toBe(true);
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
		const argStart = src.indexOf('MsgBox(') + 'MsgBox('.length;
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Prompt',
			edit: {
				span: {
					start: argStart,
					end: argStart,
				},
				newText: 'TODO_Prompt',
			},
		});
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('flags a parenless runtime function statement that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MsgBox');
		expect(hits[0].message).toContain('got 0');
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Prompt',
			edit: {
				span: {
					start: src.indexOf('MsgBox') + 'MsgBox'.length,
					end: src.indexOf('MsgBox') + 'MsgBox'.length,
				},
				newText: ' TODO_Prompt',
			},
		});
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
		expect(spanText(src, hits[0])).toBe(',');
		expect(hits[0].message).toContain('TaxRate');
		expect(hits[0].message).toContain('Argument not optional');
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'TaxRate',
			edit: {
				span: {
					start: src.indexOf(', )'),
					end: src.indexOf(', )') + 2,
				},
				newText: ', TODO_TaxRate',
			},
		});
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
		expect(spanText(src, hits[0])).toBe(',');
		expect(hits[0].message).toContain('Subtotal');
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Subtotal',
			edit: {
				span: {
					start: src.indexOf('(,') + 1,
					end: src.indexOf('(,') + 3,
				},
				newText: 'TODO_Subtotal, ',
			},
		});
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

	it('treats CallByName trailing args as a runtime ParamArray', () => {
		const src =
			'Sub Main()\n' +
			'    CallByName target, "Run", VbMethod, a0, a1, a2, a3\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('treats VBA Array ArgList as a runtime ParamArray', () => {
		const src =
			'Sub Main()\n' +
			'    Dim monthNames As Variant\n' +
			'    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _\n' +
			'                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", _\n' +
			'                       "Jan+1", "Feb+1", "Mar+1", "Apr+1", "May+1", _\n' +
			'                       "Jun+1", "Jul+1", "Aug+1")\n' +
			'    monthNames = Array()\n' +
			'End Sub\n';
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
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helper },
			], 'Caller'),
			'argument-count',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('PrintTotal');
		expect(hits[0].message).toContain('expected 2 arguments');
	});

	it('validates same-module Declare argument counts', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Sleep');
		expect(hits[0].message).toContain('expected 1 argument');
	});

	it('uses the active conditional Declare signature for same-module calls', () => {
		const src =
			'#If VBA7 Then\n' +
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'#Else\n' +
			'Private Declare Sub Sleep Lib "kernel32" ()\n' +
			'#End If\n' +
			'Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const vba7Hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: true } },
			}),
			'argument-count',
		);
		expect(vba7Hits).toHaveLength(1);
		expect(spanText(src, vba7Hits[0])).toBe('Sleep');

		const legacyHits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: false } },
			}),
			'argument-count',
		);
		expect(legacyHits).toHaveLength(0);
	});

	it('uses exported project Declare signatures for cross-module argument count', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const nativeApi =
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'NativeApi', source: nativeApi },
				]),
			}),
			'argument-count',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Sleep');
		expect(hits[0].message).toContain('expected 1 argument');
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

	it('does not treat zero-argument property result indexing as member-call arity', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    Dim child As Object\n' +
			'    Set child = p.Children(1)\n' +
			'End Sub\n';
		const person =
			'Public Property Get Children() As Collection\n' +
			'End Property\n';
		expect(
			byCode(
				analyzeModule(caller, {
					projectClassMembers: projectClassMembers([
						{ moduleName: 'Person', moduleKind: 'class', source: person },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
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

	it('honours ParamArray on source-backed class member calls', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim cb As Callback\n' +
			'    cb.Run first, second, third\n' +
			'End Sub\n';
		const callback =
			'Public Function Run(ParamArray params() As Variant) As Variant\n' +
			'End Function\n';

		expect(
			byCode(
				analyzeModule(caller, {
					projectClassMembers: projectClassMembers([
						{ moduleName: 'Callback', moduleKind: 'class', source: callback },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('validates parenless source-backed class member call statements', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.Save\n' +
			'    p.Save "ok"\n' +
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
		expect(hits[0].message).toContain('got 0');
	});

	it('validates parenless generated host member call statements', () => {
		const src =
			'Sub Main()\n' +
			'    ActiveSheet.Range\n' +
			'    ActiveSheet.Range "A1"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Range');
		expect(hits[0].message).toContain('got 0');
	});

	it('validates required arguments on runtime object member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Err.Raise\n' +
			'    Err.Raise vbObjectError + 1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Raise');
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
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('Subtotal');
		expect(hits[0].message).toContain("will raise Run-time error '13'");
	});

	it('validates same-module Declare argument types', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Public Sub T()\n' +
			'    Sleep "bad"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('Milliseconds');
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

	it('errors on decimal literals outside Byte and Integer parameter bounds', () => {
		const src =
			'Public Sub TakesByte(ByVal value As Byte)\n' +
			'End Sub\n' +
			'Public Sub TakesInteger(ByVal value As Integer)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    TakesByte 0\n' +
			'    TakesByte 255\n' +
			'    TakesByte 256\n' +
			'    TakesByte -1\n' +
			'    TakesInteger -32768\n' +
			'    TakesInteger 32767\n' +
			'    TakesInteger 32768\n' +
			'    TakesInteger -32769\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(4);
		expect(spanText(src, hits[0])).toBe('256');
		expect(hits[0].message).toContain('Byte');
		expect(hits[0].message).toContain("Run-time error '6'");
		expect(spanText(src, hits[1])).toBe('-1');
		expect(hits[1].message).toContain('Byte');
		expect(hits[1].message).toContain("Run-time error '6'");
		expect(spanText(src, hits[2])).toBe('32768');
		expect(hits[2].message).toContain('Integer');
		expect(hits[2].message).toContain("Run-time error '6'");
		expect(spanText(src, hits[3])).toBe('-32769');
		expect(hits[3].message).toContain('Integer');
		expect(hits[3].message).toContain("Run-time error '6'");
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

	it('flags known scalar variable type mismatches for ByRef parameters', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'byref-argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('amount');
		expect(hits[0].message).toContain('expects Long');
		expect(hits[0].message).toContain('declared as Integer');
	});

	it('treats omitted parameter passing markers as default ByRef', () => {
		const src =
			'Public Sub Mutate(value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(1);
	});

	it('does not apply ByRef exactness to ByVal parameters, literals, or parenthesized expressions', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub ReadValue(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Dim longAmount As Long\n' +
			'    ReadValue amount\n' +
			'    Mutate 1\n' +
			'    Mutate (longAmount)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(0);
	});

	it('uses unique exported project signatures for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		const helpers = 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helpers },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('amount');
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

	it('flags negative literal values for selected native VBA runtime argument bounds', () => {
		const src =
			'Sub T()\n' +
			'    a = Left$("abcdef", -1)\n' +
			'    b = Left("abcdef", -2)\n' +
			'    c = String$(-3, "x")\n' +
			'    d = String(-4, "x")\n' +
			'    e = VBA.Left$("abcdef", -5)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');
		expect(hits).toHaveLength(5);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['-1', '-2', '-3', '-4', '-5']);
		expect(hits[0].message).toContain('Left$');
		expect(hits[0].message).toContain('Length');
		expect(hits[0].message).toContain("Run-time error '5'");
		expect(hits[2].message).toContain('String$');
		expect(hits[2].message).toContain('Number');
	});

	it('flags oracle-backed runtime argument bounds for Right Space and Mid', () => {
		const src =
			'Sub T()\n' +
			'    a = Right$("abcdef", -1)\n' +
			'    b = Right("abcdef", -2)\n' +
			'    c = Space$(-3)\n' +
			'    d = Space(-4)\n' +
			'    e = Mid$("abcdef", 0, 1)\n' +
			'    f = Mid("abcdef", -1, 1)\n' +
			'    g = Mid$("abcdef", 1, -5)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');
		expect(hits).toHaveLength(7);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['-1', '-2', '-3', '-4', '0', '-1', '-5']);
		expect(hits[0].message).toContain('Right$');
		expect(hits[0].message).toContain('Length');
		expect(hits[2].message).toContain('Space$');
		expect(hits[2].message).toContain('Number');
		expect(hits[4].message).toContain('Mid$');
		expect(hits[4].message).toContain('Start');
		expect(hits[6].message).toContain('Length');
	});

	it('flags oracle-backed runtime argument bounds for Replace', () => {
		const src =
			'Sub T()\n' +
			'    a = Replace("abcdef", "a", "z", 0)\n' +
			'    b = Replace("abcdef", "a", "z", -1)\n' +
			'    c = Replace("aaaa", "a", "z", 1, -2)\n' +
			'    d = Replace("aaaa", "a", "z", Count:=-2, Start:=1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');
		expect(hits).toHaveLength(4);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['0', '-1', '-2', '-2']);
		expect(hits[0].message).toContain('Replace');
		expect(hits[0].message).toContain('Start');
		expect(hits[2].message).toContain('Count');
	});

	it('folds deterministic Const Enum and integer expressions for runtime argument bounds', () => {
		const src =
			'Private Const BadLength As Long = -1\n' +
			'Private Const GoodLength As Long = 0\n' +
			'Private Enum RuntimeArgStart\n' +
			'    EnumBadStart = 0\n' +
			'End Enum\n' +
			'Sub T()\n' +
			'    Const BadStart As Long = 1 - 1\n' +
			'    Const BadCount As Long = -1 - 1\n' +
			'    Const GoodCount As Long = 0 - 1\n' +
			'    a = Left$("abcdef", BadLength)\n' +
			'    b = Left$("abcdef", GoodLength)\n' +
			'    c = Left$("abcdef", 1 - 2)\n' +
			'    d = Right$("abcdef", 1 - 1)\n' +
			'    e = Replace("abcdef", "a", "z", BadStart)\n' +
			'    f = Replace("aaaa", "a", "z", 1, BadCount)\n' +
			'    g = Replace("aaaa", "a", "z", 1, GoodCount)\n' +
			'    h = Mid$("abcdef", EnumBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');
		expect(hits).toHaveLength(5);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'BadLength',
			'1 - 2',
			'BadStart',
			'BadCount',
			'EnumBadStart',
		]);
		expect(hits.map((hit) => hit.message)).toEqual([
			expect.stringContaining('is -1'),
			expect.stringContaining('is -1'),
			expect.stringContaining('is 0'),
			expect.stringContaining('is -2'),
			expect.stringContaining('is 0'),
		]);
	});

	it('flags oracle-backed InStr Start values below one without flagging two-argument calls', () => {
		const src =
			'Sub T()\n' +
			'    Const BadStart As Long = 0\n' +
			'    a = InStr(0, "abcdef", "a")\n' +
			'    b = InStr(-1, "abcdef", "a")\n' +
			'    c = InStr(1, "abcdef", "a")\n' +
			'    d = InStr("abcdef", "a")\n' +
			'    e = InStr(BadStart, "abcdef", "a")\n' +
			'    f = InStr(Start:=0, String1:="abcdef", String2:="a")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');

		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['0', '-1', 'BadStart']);
		expect(hits.map((hit) => hit.message)).toEqual([
			expect.stringContaining("'Start' of 'InStr' is 0"),
			expect.stringContaining("'Start' of 'InStr' is -1"),
			expect.stringContaining("'Start' of 'InStr' is 0"),
		]);
	});

	it('flags oracle-backed Chr and ChrW CharCode values outside proven bounds', () => {
		const src =
			'Sub T()\n' +
			'    Const BadChrLow As Long = -1\n' +
			'    Const BadChrHigh As Long = 256\n' +
			'    Const BadChrWHigh As Long = 65536\n' +
			'    a = Chr(-1)\n' +
			'    b = Chr(0)\n' +
			'    c = Chr(255)\n' +
			'    d = Chr(256)\n' +
			'    e = Chr(BadChrLow)\n' +
			'    f = Chr(BadChrHigh)\n' +
			'    g = ChrW(-1)\n' +
			'    h = ChrW(65535)\n' +
			'    i = ChrW(65536)\n' +
			'    j = ChrW(BadChrWHigh)\n' +
			'    k = VBA.Chr(256)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');

		expect(hits).toHaveLength(7);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'-1',
			'256',
			'BadChrLow',
			'BadChrHigh',
			'65536',
			'BadChrWHigh',
			'256',
		]);
		expect(hits.map((hit) => hit.message)).toEqual([
			expect.stringContaining("'CharCode' of 'Chr' is -1"),
			expect.stringContaining("'CharCode' of 'Chr' is 256"),
			expect.stringContaining("'CharCode' of 'Chr' is -1"),
			expect.stringContaining("'CharCode' of 'Chr' is 256"),
			expect.stringContaining("'CharCode' of 'ChrW' is 65536"),
			expect.stringContaining("'CharCode' of 'ChrW' is 65536"),
			expect.stringContaining("'CharCode' of 'Chr' is 256"),
		]);
	});

	it('folds visible cross-module Const and Enum values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Left$("abcdef", SharedGoodLength)\n' +
			'    c = Mid$("abcdef", SharedBadStart, 1)\n' +
			'    d = Replace("aaaa", "a", "z", 1, SharedBadCount)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Const SharedGoodLength As Long = 0\n' +
						'Public Const SharedBadCount As Long = -2\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedBadLength',
			'SharedBadStart',
			'SharedBadCount',
		]);
	});

	it('folds module-qualified cross-module Const and Enum values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedRuntimeArgs.SharedBadLength)\n' +
			'    b = Left$("abcdef", SharedRuntimeArgs.SharedGoodLength)\n' +
			'    c = Mid$("abcdef", SharedRuntimeArgs.SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Const SharedGoodLength As Long = 0\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedRuntimeArgs.SharedBadLength',
			'SharedRuntimeArgs.SharedBadStart',
		]);
	});

	it('folds exported cross-module Const values through private same-module helpers for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedBadLengthFromHidden)\n' +
			'    b = Left$("abcdef", SharedGoodLengthFromHidden)\n' +
			'    c = Left$("abcdef", SharedRuntimeArgs.SharedBadLengthFromHidden)\n' +
			'    d = Left$("abcdef", SharedRuntimeArgs.SharedGoodLengthFromHidden)\n' +
			'    e = Left$("abcdef", SharedRuntimeArgs.HiddenBadLength)\n' +
			'End Sub\n';
		const shared =
			'Private Const HiddenBadLength As Long = -1\n' +
			'Private Const HiddenGoodLength As Long = 0\n' +
			'Public Const SharedBadLengthFromHidden As Long = HiddenBadLength\n' +
			'Public Const SharedGoodLengthFromHidden As Long = HiddenGoodLength\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'SharedRuntimeArgs', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'runtime-argument-value',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedBadLengthFromHidden',
			'SharedRuntimeArgs.SharedBadLengthFromHidden',
		]);
	});

	it('defers hidden and ambiguous cross-module Const values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", HiddenBadLength)\n' +
			'    b = Left$("abcdef", AmbiguousLength)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'PrivateGlobals',
					source:
						'Private Const HiddenBadLength As Long = -1\n' +
						'Public Const AmbiguousLength As Long = -1\n',
				},
				{
					moduleName: 'MoreGlobals',
					source: 'Public Const AmbiguousLength As Long = 0\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(0);
	});

	it('lets current-module and local Const values shadow cross-module runtime argument bounds', () => {
		const caller =
			'Private Const SharedBadLength As Long = 0\n' +
			'Sub T()\n' +
			'    Const SharedBadStart As Long = 1\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Mid$("abcdef", SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(0);
	});

	it('accepts zero and unknown runtime argument values for selected native bounds', () => {
		const src =
			'Sub T()\n' +
			'    Dim count As Long\n' +
			'    a = Left$("abcdef", 0)\n' +
			'    b = Left("abcdef", count)\n' +
			'    c = String$(0, "x")\n' +
			'    d = String(count, "x")\n' +
			'    e = Right$("abcdef", 0)\n' +
			'    f = Right("abcdef", count)\n' +
			'    g = Space$(0)\n' +
			'    h = Space(count)\n' +
			'    i = Mid$("abcdef", 1, 0)\n' +
			'    j = Mid("abcdef", count, count)\n' +
			'    k = Replace("abcdef", "a", "z", 1)\n' +
			'    l = Replace("aaaa", "a", "z", 1, -1)\n' +
			'    m = Replace("aaaa", "a", "z", 1, 0)\n' +
			'    n = Replace("aaaa", "a", "z", count, count)\n' +
			'    o = Replace$("aaaa", "a", "z", 0)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'runtime-argument-value')).toHaveLength(0);
	});

	it('does not treat shadowed Left calls as native runtime argument-value checks', () => {
		const src =
			'Public Function Left(ByVal text As String, ByVal count As Long) As String\n' +
			'End Function\n' +
			'Sub T()\n' +
			'    Dim localValue As Long\n' +
			'    Dim Left As Long\n' +
			'    a = Left("abcdef", -1)\n' +
			'    b = VBA.Left$("abcdef", -1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('-1');
		expect(hits[0].message).toContain('Left$');
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

	it('uses parenless source-backed class member signatures for argument types', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.Save "bad"\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Count As Long)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('Count');
		expect(hits[0].message).toContain('Long');
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

	it('errors on decimal literals outside Byte and Integer assignment bounds', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim small As Byte\n' +
			'    Dim count As Integer\n' +
			'    small = 0\n' +
			'    small = 255\n' +
			'    small = 256\n' +
			'    small = -1\n' +
			'    count = -32768\n' +
			'    count = 32767\n' +
			'    count = 32768\n' +
			'    count = -32769\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');
		expect(hits).toHaveLength(4);
		expect(spanText(src, hits[0])).toBe('256');
		expect(hits[0].message).toContain('Byte');
		expect(hits[0].message).toContain("Run-time error '6'");
		expect(spanText(src, hits[1])).toBe('-1');
		expect(hits[1].message).toContain('Byte');
		expect(hits[1].message).toContain("Run-time error '6'");
		expect(spanText(src, hits[2])).toBe('32768');
		expect(hits[2].message).toContain('Integer');
		expect(hits[2].message).toContain("Run-time error '6'");
		expect(spanText(src, hits[3])).toBe('-32769');
		expect(hits[3].message).toContain('Integer');
		expect(hits[3].message).toContain("Run-time error '6'");
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

	it('uses visible UDT fields for missing-member and assignment type diagnostics', () => {
		const types = 'Public Type TPoint\n    X As Long\nEnd Type\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As TPoint\n' +
			'    p.X = "bad"\n' +
			'    p.Missing = 1\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			moduleName: 'Caller',
			projectClassMembers: projectMemberSurfaces(
				[
					{ moduleName: 'Caller', source: src },
					{ moduleName: 'Types', source: types },
				],
				'Caller',
			),
		});
		const typeHits = byCode(diagnostics, 'assignment-type-mismatch');
		expect(typeHits).toHaveLength(1);
		expect(spanText(src, typeHits[0])).toBe('"bad"');
		expect(typeHits[0].message).toContain('p.X');
		const memberHits = byCode(diagnostics, 'member-not-found');
		expect(memberHits).toHaveLength(1);
		expect(spanText(src, memberHits[0])).toBe('Missing');
		expect(memberHits[0].message).toContain('TPoint.Missing');
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

	it('errors when a known standard-module qualifier uses an unknown exported member', () => {
		const caller =
			'Public Sub T()\n' +
			'    Globals.SharedValue = 1\n' +
			'    Globals.MissingValue = 2\n' +
			'    Globals.MissingProcedure\n' +
			'End Sub\n';
		const globals =
			'Public SharedValue As Long\n' +
			'Public Sub Save()\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', moduleKind: 'standard', source: globals },
		], 'Caller');
		const hits = byCode(diagnostics, 'member-not-found');
		expect(hits.map((hit) => spanText(caller, hit))).toEqual(['MissingValue', 'MissingProcedure']);
		expect(hits[0].message).toContain('Globals.MissingValue');
		expect(hits[1].message).toContain('Globals.MissingProcedure');
	});

	it('does not expose private standard-module members through qualified member diagnostics', () => {
		const caller =
			'Public Sub T()\n' +
			'    Helpers.Hidden\n' +
			'End Sub\n';
		const helpers =
			'Private Sub Hidden()\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', moduleKind: 'standard', source: helpers },
		], 'Caller');
		const hits = byCode(diagnostics, 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Hidden');
		expect(hits[0].message).toContain('Helpers.Hidden');
	});

	it('does not treat unknown external-style qualifiers as project member surfaces', () => {
		const caller =
			'Public Sub T()\n' +
			'    Scripting.Dictionary.CompareMode = TextCompare\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [], 'Caller');
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
			'    ActiveSheet.Buttons\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdf');
		expect(hits[0].message).toContain('Excel.Worksheet.asdf');
	});

	it('uses the exhaustive Worksheet host surface through workbook worksheet chains', () => {
		const src =
			'Public Sub T()\n' +
			'    Workbooks(1).Worksheets(1).asdf\n' +
			'    Workbooks(1).Worksheets(1).Range("A1")\n' +
			'    Workbooks(1).Worksheets(1).Buttons\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdf');
		expect(hits[0].message).toContain('Excel.Worksheet.asdf');
	});

	it('uses the exhaustive Range host surface for declared, global, and chained receivers', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim rng As Range\n' +
			'    ActiveCell.NoSuchMember\n' +
			'    ActiveCell.Value2\n' +
			'    rng.DoesNotExist\n' +
			'    rng.ClearContents\n' +
			'    Workbooks(1).Worksheets(1).Range("A1").MissingRangeMember\n' +
			'    Workbooks(1).Worksheets(1).Range("A1").Offset(1, 0).Value = 1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'NoSuchMember',
			'DoesNotExist',
			'MissingRangeMember',
		]);
		expect(hits.every((hit) => hit.message.includes('Excel.Range.'))).toBe(true);
	});

	it('uses the exhaustive runtime object surface for Err', () => {
		const src =
			'Public Sub T()\n' +
			'    Err.Raise vbObjectError + 1, "M", "boom"\n' +
			'    Err.DoesNotExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
		const hits = byCode(diagnostics, 'member-not-found');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
		expect(hits[0].message).toContain('Err.DoesNotExist');
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
			'    ws.Buttons\n' +
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
			'Public Function myFunction()\n' +
			'\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('warning');
		expect(spanText(src, hits[0])).toBe('myFunction');
		expect(hits[0].message).toContain('default value');
	});

	it('does not warn when a typed Function or Property Get falls through', () => {
		const src =
			'Public Function Label() As String\n' +
			'End Function\n' +
			'\n' +
			'Public Property Get Name() As String\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
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

	it('accepts Function return variables passed to known ByRef helper parameters', () => {
		const src =
			'Public Function Run(ParamArray params() As Variant)\n' +
			'    Call CopyVariant(Run, RunEx(params))\n' +
			'End Function\n' +
			'\n' +
			'Private Function RunEx(ByVal values As Variant) As Variant\n' +
			'    RunEx = values\n' +
			'End Function\n' +
			'\n' +
			'Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to project-visible ByRef helper parameters', () => {
		const runSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    Call CopyVariant(dest:=Run, value:=1)\n' +
			'End Function\n';
		const helpersSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			runSrc,
			[
				{ moduleName: 'Runner', source: runSrc },
				{ moduleName: 'Helpers', source: helpersSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to module-qualified ByRef helper parameters', () => {
		const runSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    Helpers.CopyVariant Run, 1\n' +
			'End Function\n';
		const helpersSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			runSrc,
			[
				{ moduleName: 'Runner', source: runSrc },
				{ moduleName: 'Helpers', source: helpersSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(0);
	});

	it('does not count ByVal or ambiguous project helper arguments as return assignments', () => {
		const byValRunSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    CopyVariant Run, 1\n' +
			'End Function\n';
		const byValHelperSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByVal dest As Variant, ByVal value As Variant)\n' +
			'End Sub\n';
		const ambiguousRunSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    CopyVariant Run, 1\n' +
			'End Function\n';
		const byRefHelperSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			byValRunSrc,
			[
				{ moduleName: 'Runner', source: byValRunSrc },
				{ moduleName: 'Helpers', source: byValHelperSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(1);
		expect(byCode(analyzeProjectModule(
			ambiguousRunSrc,
			[
				{ moduleName: 'Runner', source: ambiguousRunSrc },
				{ moduleName: 'HelpersA', source: byRefHelperSrc },
				{ moduleName: 'HelpersB', source: byRefHelperSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(1);
	});

	it('accepts return assignments in the active default VBA7 branch', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			'    HandleValue = 7\n' +
			'#Else\n' +
			'    HandleValue = 6\n' +
			'#End If\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts active VBA7 return assignments with blank lines before #End If', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			' HandleValue = 0\n' +
			'#Else\n' +
			'HandleValue = 1\n' +
			'\n' +
			'\n' +
			'#End If\n' +
			'    \n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('does not count return assignments from inactive default VBA7 branches', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			'    Debug.Print "active branch"\n' +
			'#Else\n' +
			'    HandleValue = 6\n' +
			'#End If\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('HandleValue');
	});

	it('does not warn on conditionally split VBA7 Function headers with a shared body', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Function HandleValue()\n' +
			'#Else\n' +
			'Public Function HandleValue()\n' +
			'#End If\n' +
			'    HandleValue = 1\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('keeps parsing after conditionally split procedure headers with a shared body', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Function Create(ByVal pointer As LongPtr) As Object\n' +
			'#Else\n' +
			'Public Function Create(ByVal pointer As Long) As Object\n' +
			'#End If\n' +
			'    Set Create = Nothing\n' +
			'End Function\n' +
			'Friend Sub Init(ByVal mode As Long)\n' +
			'    Select Case mode\n' +
			'        Case 1\n' +
			'            With Create(0)\n' +
			'                .Name = "ready"\n' +
			'            End With\n' +
			'    End Select\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('does not flag unknown-constant alternate procedure headers as declarations in a procedure', () => {
		const src =
			'#If FULL_INTELLISENSE Then\n' +
			'Public Function AsAcc() As stdAcc\n' +
			'#Else\n' +
			'Public Function AsAcc() As Object\n' +
			'#End If\n' +
			'    Set AsAcc = Nothing\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('ignores exported member Attribute lines while parsing procedure bodies', () => {
		const src =
			'Friend Sub Init(ByVal mode As Long)\n' +
			'Attribute Init.VB_Description = "Initialises this object."\n' +
			'    Select Case mode\n' +
			'        Case 1\n' +
			'            With Nothing\n' +
			'                .Name = "ready"\n' +
			'            End With\n' +
			'    End Select\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('checks Property Get procedures and ignores Subs', () => {
		const src =
			'Public Property Get Name()\n' +
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

describe('analyzeModule - division by zero', () => {
	it('errors on literal zero divisors for division, integer division, and Mod', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'    a = 1 \\ 0\n' +
			'    a = 1 Mod 0\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['0', '0', '0']);
		expect(hits[0].message).toContain("Run-time error '11'");
	});

	it('detects parenthesized and signed zero divisors inside nested expressions', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Variant\n' +
			'    a = IIf(True, 1, 1 / (-0))\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('-0');
	});

	it('errors on non-decimal zero literal divisors', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Long\n' +
			'    a = 1 / &H0\n' +
			'    a = 1 \\ &O0\n' +
			'    a = 1 Mod &H1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['&H0', '&O0']);
	});

	it('errors on hex and octal Const zero divisors', () => {
		const src =
			'Private Const ModuleHexZero As Long = &H0\n' +
			'Public Sub T()\n' +
			'    Const LocalOctalZero As Long = &O0\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ModuleHexZero\n' +
			'    a = 1 \\ LocalOctalZero\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleHexZero', 'LocalOctalZero']);
	});

	it('folds hex Const expressions for zero divisor checks', () => {
		const src =
			'Public Sub T()\n' +
			'    Const HexZero As Long = &H1 - &H1\n' +
			'    Const HexOne As Long = &H1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / HexZero\n' +
			'    a = 1 Mod HexOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('HexZero');
	});

	it('errors on zero-valued Enum member divisors', () => {
		const src =
			'Private Enum Denominator\n' +
			'    ExplicitZero = 0\n' +
			'    ImplicitOne\n' +
			'    ExplicitTwo = 2\n' +
			'End Enum\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ExplicitZero\n' +
			'    a = 1 / ImplicitOne\n' +
			'    a = 1 Mod ExplicitTwo\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ExplicitZero');
	});

	it('folds implicit and expression Enum member values for zero divisor checks', () => {
		const src =
			'Public Enum Denominator\n' +
			'    ImplicitZero\n' +
			'    ImplicitOne\n' +
			'    ExpressionZero = ImplicitOne - 1\n' +
			'End Enum\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 \\ ImplicitZero\n' +
			'    a = 1 / ExpressionZero\n' +
			'    a = 1 Mod ImplicitOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ImplicitZero', 'ExpressionZero']);
	});

	it('errors on zero-valued module and local Const divisors', () => {
		const src =
			'Private Const ModuleZero As Long = 0\n' +
			'Private Const ModuleOne As Long = 0 + 1\n' +
			'Public Sub T()\n' +
			'    Const LocalZero As Long = 1 - 1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ModuleZero\n' +
			'    a = 1 / LocalZero\n' +
			'    a = 1 / ModuleOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleZero', 'LocalZero']);
	});

	it('folds visible cross-module Const and Enum values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 / SharedOne\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'    a = 1 \\ SharedOneDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Const SharedOne As Long = 1\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'    SharedOneDivisor = 1\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual(['SharedZero', 'SharedZeroDivisor']);
	});

	it('folds module-qualified cross-module Const and Enum values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedDivisors.SharedZero\n' +
			'    a = 1 / SharedDivisors.SharedOne\n' +
			'    a = 1 \\ SharedDivisors.SharedZeroDivisor\n' +
			'    a = 1 \\ SharedDivisors.SharedOneDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Const SharedOne As Long = 1\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'    SharedOneDivisor = 1\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedDivisors.SharedZero',
			'SharedDivisors.SharedZeroDivisor',
		]);
	});

	it('folds exported cross-module Const values through private same-module helpers for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZeroFromHidden\n' +
			'    a = 1 / SharedOneFromHidden\n' +
			'    a = 1 / SharedDivisors.SharedZeroFromHidden\n' +
			'    a = 1 / SharedDivisors.SharedOneFromHidden\n' +
			'    a = 1 / SharedDivisors.HiddenZero\n' +
			'End Sub\n';
		const shared =
			'Private Const HiddenZero As Long = 0\n' +
			'Private Const HiddenOne As Long = 1\n' +
			'Public Const SharedZeroFromHidden As Long = HiddenZero\n' +
			'Public Const SharedOneFromHidden As Long = HiddenOne\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedZeroFromHidden',
			'SharedDivisors.SharedZeroFromHidden',
		]);
	});

	it('defers hidden and ambiguous cross-module Const values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / HiddenZero\n' +
			'    a = 1 / AmbiguousZero\n' +
			'End Sub\n';
		const hidden =
			'Private Const HiddenZero As Long = 0\n';
		const first =
			'Public Const AmbiguousZero As Long = 0\n';
		const second =
			'Public Const AmbiguousZero As Long = 0\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'HiddenDivisors', source: hidden, moduleKind: 'standard' },
				{ moduleName: 'FirstDivisors', source: first, moduleKind: 'standard' },
				{ moduleName: 'SecondDivisors', source: second, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('lets current-module and local Const values shadow cross-module zero divisors', () => {
		const caller =
			'Private Const SharedZero As Long = 1\n' +
			'Public Sub T()\n' +
			'    Const SharedZeroDivisor As Long = 1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('folds parenthesized Const expressions used as divisors', () => {
		const src =
			'Private Const Zero As Long = 0\n' +
			'Private Const One As Long = 1\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / (Zero + 0)\n' +
			'    a = 1 / (Zero + One)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Zero + 0');
	});

	it('does not flag nonzero literals, variables, or nonzero parenthesized expressions', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    Dim denominator As Double\n' +
			'    a = 1 / 2\n' +
			'    a = 1 / denominator\n' +
			'    a = 1 / (0 + 1)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(0);
	});

	it('ignores literal zero divisors in inactive conditional-compilation branches', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'#If Enabled Then\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, {
				conditionalCompilation: { constants: { Enabled: false } },
			}), 'division-by-zero'),
		).toHaveLength(0);
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
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'Types', source: 'Public Enum Status\n    Active\nEnd Enum\n' },
		];

		expect(
			byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-as-type-name'),
		).toHaveLength(0);
	});

	it('accepts OLE Automation interfaces even when project enum members share the name', () => {
		const src =
			'Public Enum EKnownIID\n' +
			'    IUnknown\n' +
			'End Enum\n' +
			'Private Declare PtrSafe Function RegisterActiveObject Lib "oleaut32" (ByVal pUnk As IUnknown) As Long\n' +
			'Private Declare PtrSafe Function RegisterQualified Lib "oleaut32" (ByVal pUnk As stdole.IUnknown) As Long\n';
		expect(byCode(analyzeModule(src), 'invalid-as-type-name')).toHaveLength(0);
	});

	it('uses the default VBA7 branch for project types defined in paired branches', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Payload\n' +
			'End Sub\n';
		const types =
			'#If VBA7 Then\n' +
			'Public Type Payload\n' +
			'    Id As LongPtr\n' +
			'End Type\n' +
			'#Else\n' +
			'Public Type Payload\n' +
			'    Id As Long\n' +
			'End Type\n' +
			'#End If\n';
		const modules = [
			{ moduleName: 'Consumer', source: src },
			{ moduleName: 'Types', source: types },
		];

		const diagnostics = analyzeModule(src, {
			moduleName: 'Consumer',
			projectTypes: visibleProjectTypes(modules, 'Consumer'),
			knownNonTypeNames: visibleProjectNonTypeNames(modules, 'Consumer'),
		});
		expect(byCode(diagnostics, 'invalid-as-type-name')).toHaveLength(0);
	});

	it('accepts qualified visible project type names in declaration positions', () => {
		const src =
			'Public Type Payload\n' +
			'    Location As Geometry.TPoint\n' +
			'End Type\n' +
			'\n' +
			'Public Sub T(ByVal state As Workflow.Status)\n' +
			'    Dim p As Geometry.TPoint\n' +
			'End Sub\n';
		const modules = [
			{
				moduleName: 'Geometry',
				source: 'Public Type TPoint\n    X As Long\nEnd Type\n',
			},
			{
				moduleName: 'Workflow',
				source: 'Public Enum Status\n    Active\nEnd Enum\n',
			},
			{
				moduleName: 'OtherGeometry',
				source: 'Public Type TPoint\n    Y As Long\nEnd Type\n',
			},
		];

		expect(
			byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-as-type-name'),
		).toHaveLength(0);
	});

	it('accepts creatable project classes and UserForms after New', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As New Person\n' +
			'    Dim f As New CustomerForm\n' +
			'    Dim value As Object\n' +
			'    Set value = New Person\n' +
			'    Set value = New CustomerForm\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Person', moduleKind: 'class' as const, source: '' },
			{ moduleName: 'CustomerForm', moduleKind: 'userform' as const, source: '' },
		];

		expect(
			byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-new-type-name'),
		).toHaveLength(0);
	});

	it('flags resolved non-creatable types after New', () => {
		const src =
			'Public Type Payload\n' +
			'    Id As Long\n' +
			'End Type\n' +
			'\n' +
			'Public Enum Status\n' +
			'    Active\n' +
			'End Enum\n' +
			'\n' +
			'Public Sub T()\n' +
			'    Dim eagerSheet As New Worksheet\n' +
			'    Dim eagerLong As New Long\n' +
			'    Dim value As Object\n' +
			'    Set value = New Status\n' +
			'    Set value = New Payload\n' +
			'    Set value = New Sheet1\n' +
			'    Set value = New Worksheet\n' +
			'    Set value = New Long\n' +
			'End Sub\n';
		const modules = [
			{ moduleName: 'Sheet1', moduleKind: 'document' as const, source: '' },
		];

		const diags = analyzeProjectModule(src, modules, 'Consumer');
		const hits = byCode(diags, 'invalid-new-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'Worksheet',
			'Long',
			'Status',
			'Payload',
			'Sheet1',
			'Worksheet',
			'Long',
		]);
		expect(byCode(diags, 'invalid-as-type-name')).toHaveLength(0);
	});

	it('flags qualified non-creatable project types after New', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim value As Object\n' +
			'    Set value = New Geometry.TPoint\n' +
			'    Set value = New Workflow.Status\n' +
			'End Sub\n';
		const modules = [
			{
				moduleName: 'Geometry',
				source: 'Public Type TPoint\n    X As Long\nEnd Type\n',
			},
			{
				moduleName: 'Workflow',
				source: 'Public Enum Status\n    Active\nEnd Enum\n',
			},
		];

		const hits = byCode(analyzeProjectModule(src, modules, 'Consumer'), 'invalid-new-type-name');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['TPoint', 'Status']);
	});

	it('defers unresolved New type names to the project-wide binder and external references', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim eager As New ExternalThing\n' +
			'    Dim value As Object\n' +
			'    Set value = New OtherExternalThing\n' +
			'End Sub\n';
		const diags = analyzeModule(src);

		expect(byCode(diags, 'invalid-new-type-name')).toHaveLength(0);
		expect(byCode(diags, 'invalid-as-type-name')).toHaveLength(0);
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

	it('accepts fixed-length string declarations and qualified type names', () => {
		const src =
			'Private fixedName As String * 10\n' +
			'Private Type Header\n' +
			'    Code As String * 4\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Dim localName As String * 20\n' +
			'    Dim workbook As Excel.Workbook\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('flags extra tokens after a fixed-length string suffix', () => {
		const src = 'Sub T()\n    Dim fixedName As String * 10 junk\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unexpected-declaration-token');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('junk');
	});
});

describe('analyzeModule - fixed-length String bounds', () => {
	it('accepts verified literal boundaries and literal Const lengths', () => {
		const src =
			'Private Const HeaderCodeLength As Long = 65526\n' +
			'Private moduleName As String * 1\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const MaxNameLength As Long = 20\n' +
			'    Dim localName As String * MaxNameLength\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-length-string-size')).toHaveLength(0);
	});

	it('resolves compound integer Const expressions used as fixed-length String sizes', () => {
		const src =
			'Private Const BaseLength As Long = 32763\n' +
			'Private Const HeaderCodeLength As Long = BaseLength * 2 + 1\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const LocalBase As Long = 10\n' +
			'    Const MaxNameLength As Long = LocalBase + 10\n' +
			'    Const LocalTooSmall As Long = LocalBase - 10\n' +
			'    Dim localName As String * MaxNameLength\n' +
			'    Dim badLocalName As String * LocalTooSmall\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['HeaderCodeLength', 'LocalTooSmall']);
		expect(hits[0].message).toContain('got 65527');
		expect(hits[1].message).toContain('got 0');
	});

	it('resolves non-decimal integer Consts used as fixed-length String sizes', () => {
		const src =
			'Private Const HeaderCodeLength As Long = &H14\n' +
			'Private Type Header\n' +
			'    Code As String * HeaderCodeLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const LocalTooSmall As Long = &O0\n' +
			'    Dim localName As String * LocalTooSmall\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('LocalTooSmall');
		expect(hits[0].message).toContain('got 0');
	});

	it('defers unknown and non-deterministic Const length expressions', () => {
		const src =
			'Private Const FractionLength As Long = 20 / 2\n' +
			'Private moduleName As String * MissingLength\n' +
			'Private Type Header\n' +
			'    Code As String * FractionLength\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Const RuntimeLength As Long = CLng(20)\n' +
			'    Dim localName As String * RuntimeLength\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-length-string-size')).toHaveLength(0);
	});

	it('flags literal lengths outside the VBE-verified range everywhere declarations are parsed', () => {
		const src =
			'Private moduleName As String * 0\n' +
			'Private Type Header\n' +
			'    Code As String * 65527\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    If True Then\n' +
			'        Dim localName As String * 0\n' +
			'    End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['0', '65527', '0']);
		expect(hits[0].message).toContain('between 1 and 65526');
		expect(byCode(analyzeModule(src), 'unexpected-declaration-token')).toHaveLength(0);
	});

	it('flags simple Const lengths outside the VBE-verified range', () => {
		const src =
			'Private Const ModuleTooLong As Long = 65527\n' +
			'Private moduleName As String * ModuleTooLong\n' +
			'Sub T()\n' +
			'    Const LocalTooSmall As Long = 0\n' +
			'    Dim localName As String * LocalTooSmall\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleTooLong', 'LocalTooSmall']);
		expect(hits[0].message).toContain('got 65527');
		expect(hits[1].message).toContain('got 0');
	});

	it('lets procedure-local Const lengths shadow module Const lengths', () => {
		const src =
			'Private Const NameLength As Long = 65527\n' +
			'Sub T()\n' +
			'    Const NameLength As Long = 20\n' +
			'    Dim localName As String * NameLength\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-length-string-size')).toHaveLength(0);
	});

	it('resolves bracketed Const names used as fixed-length String sizes', () => {
		const src =
			'Private Const [Name Length] As Long = 0\n' +
			'Sub T()\n' +
			'    Dim localName As String * [Name Length]\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-length-string-size');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('[Name Length]');
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

	it('accepts private fixed-length strings across object module kinds', () => {
		const src = 'Private FixedName As String * 20\n';
		for (const moduleKind of ['class', 'document', 'userform'] as const) {
			const diagnostics = analyzeModule(src, { moduleName: 'ObjectModule', moduleKind });
			expect(byCode(diagnostics, 'object-module-public-member')).toHaveLength(0);
			expect(byCode(diagnostics, 'fixed-length-string-size')).toHaveLength(0);
		}
	});
});

describe('analyzeModule - Event declaration module-kind restrictions', () => {
	it('flags Event declarations in standard modules', () => {
		const src =
			'Public Event BeforeAdd(ByRef arr As Variant, ByRef cancel As Boolean)\n' +
			'Private Event AfterAdd(ByRef arr As Variant)\n';
		const hits = byCode(analyzeModule(src), 'event-declaration-module-kind');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['BeforeAdd', 'AfterAdd']);
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('class, document, or UserForm modules');
	});

	it('ignores inactive standard-module Event declarations', () => {
		const src =
			'#If VBA7 Then\n' +
			'#Else\n' +
			'Public Event LegacyOnly()\n' +
			'#End If\n';
		expect(byCode(analyzeModule(src), 'event-declaration-module-kind')).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'event-declaration-module-kind',
			),
		).toHaveLength(1);
	});
});

describe('analyzeModule - conditional Declare platform rules', () => {
	it('requires PtrSafe only when the supplied compiler constants prove Win64', () => {
		const src = 'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';

		const hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { Win64: true } },
			}),
			'declare-missing-ptrsafe',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Sleep');

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { Win64: false } },
				}),
				'declare-missing-ptrsafe',
			),
		).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'declare-missing-ptrsafe')).toHaveLength(0);
	});

	it('does not flag inactive legacy Declare branches under Win64', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'#Else\n' +
			'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
			'#End If\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true, Win64: true } },
				}),
				'declare-missing-ptrsafe',
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

	it('does not guide for chart handlers in chart document modules', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Chart_Calculate()\nEnd Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					moduleName: 'RevenueChart',
					moduleKind: 'document',
					documentType: 'chart',
				}),
				'event-handler-module-scope',
			),
		).toHaveLength(0);
	});

	it('guides when a chart handler is declared in a worksheet document module', () => {
		const src =
			'Option Explicit\n' +
			'Private Sub Chart_Calculate()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Sheet1',
				moduleKind: 'document',
				documentType: 'worksheet',
			}),
			'event-handler-module-scope',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Chart_Calculate');
		expect(hits[0].message).toContain('worksheet document module');
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

	it('leaves required-argument procedure empty calls to arity diagnostics', () => {
		const src =
			'Sub mySub()\n' +
			'    Greet()\n' +
			'End Sub\n' +
			'\n' +
			'Sub Greet(ByVal message As String)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Greet');
	});

	it('flags a standalone zero-argument runtime call with empty parentheses', () => {
		const src = 'Sub T()\n    DoEvents()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoEvents()');
		expect(hits[0].message).toContain("use 'DoEvents' as a statement");
		expect(hits[0].message).not.toContain('prefixed with Call');
	});

	it('flags DoEvents as an invalid explicit Call target', () => {
		const src =
			'Sub T()\n' +
			'    Call DoEvents\n' +
			'    Call DoEvents()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-explicit-call-target');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DoEvents', 'DoEvents']);
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
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
		const hits = byCode(
			analyzeProjectModule(src, [
				{ moduleName: 'Helpers', source: 'Public Function myFunction() As String\nEnd Function\n' },
			], 'Caller'),
			'call-statement-forbids-parens',
		);
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

	it('accepts zero-argument runtime calls as bare statements or in expression context', () => {
		const src =
			'Sub T()\n' +
			'    Dim value As Integer\n' +
			'    DoEvents\n' +
			'    value = DoEvents()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'invalid-explicit-call-target')).toHaveLength(0);
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

	it('flags a unique exported project Function called with parenless arguments in an assignment', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'expression-call-requires-parens',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('InvoiceTotal');
	});

	it('flags a module-qualified project Function called with parenless arguments in an assignment', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = Invoices.InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'expression-call-requires-parens',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('InvoiceTotal');
	});

	it('does not guess when a bare exported project Function name is ambiguous', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const first =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const second =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n';

		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'Invoices', source: first },
			{ moduleName: 'AlternateInvoices', source: second },
		], 'Caller'), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('does not treat exported project Subs as expression Functions', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = PrintTotal 100\n' +
			'End Sub\n';
		const helpers = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';

		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', source: helpers },
		], 'Caller'), 'expression-call-requires-parens')).toHaveLength(0);
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

	it('accepts parameterless property reads followed by keyword operators', () => {
		const src =
			'Public Property Get Style() As Long\nEnd Property\n' +
			'Public Sub T()\n' +
			'    Dim resizable As Boolean\n' +
			'    resizable = Style And &H40000\n' +
			'    resizable = Style Or &H40000\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid expression syntax', () => {
	it('flags an impossible operator sequence in a call argument expression', () => {
		const src =
			'Sub T()\n' +
			'    MsgBox myFunctionTest***\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('***');
	});

	it('flags unsupported C-style ternary syntax', () => {
		const src = 'Sub T()\n    value = flag ? 1 : 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('?');
		expect(hits[0].message).toContain("'?' conditional operator");
	});

	it('flags a statement that ends with a binary operator', () => {
		const src = 'Sub T()\n    total = subtotal *\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('*');
	});

	it('flags trailing member-access dots on object receivers', () => {
		const src = 'Sub T()\n    ThisWorkbook.\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
		expect(hits[0].message).toContain('member name');
	});

	it('flags a bare leading dot inside With as incomplete final source', () => {
		const src =
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('accepts complete leading-dot member access inside With', () => {
		const src =
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .Range("A1").Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('lets scalar member access own known scalar trailing dots', () => {
		const src = 'Sub T()\n    Dim value As String\n    value.\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(1);
	});

	it('treats fixed-length String declarations as scalar String receivers', () => {
		const src = 'Sub T()\n    Dim value As String * 20\n    value.\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
	});

	it('does not flag valid arithmetic or string expressions', () => {
		const src =
			'Sub T()\n' +
			'    total = subtotal * taxRate\n' +
			'    message = prefix & suffix\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('accepts IIf as the supported inline conditional function', () => {
		const src = 'Sub T()\n    value = IIf(flag, 1, 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});
});

describe('analyzeModule - parameter order', () => {
	it('flags a required parameter after an Optional one', () => {
		const src =
			'Sub T(Optional ByVal a As Long = 1, ByVal b As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'required-param-after-optional');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('b');
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
		expect(spanText(src, hits[0])).toBe('items');
	});

	it('accepts a trailing ParamArray', () => {
		const src = 'Sub T(ByVal n As Long, ParamArray items() As Variant)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'paramarray-not-last')).toHaveLength(0);
	});
});

describe('analyzeModule - parameter default values', () => {
	it('flags nonnumeric string defaults for numeric and Boolean Optional parameters', () => {
		const src =
			'Sub T(Optional ByVal count As Long = "bad", Optional ByVal enabled As Boolean = "bad")\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'parameter-default-type-mismatch');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['"bad"', '"bad"']);
		expect(hits[0].message).toContain('count');
		expect(hits[0].message).toContain('expects Long');
		expect(hits[1].message).toContain('enabled');
		expect(hits[1].message).toContain('expects Boolean');
	});

	it('accepts oracle-backed scalar Optional default controls', () => {
		const src =
			'Sub T(Optional ByVal count As Long = 1, Optional ByVal fromText As Long = "1", Optional ByVal label As String = "ok", Optional ByVal enabled As Boolean = True)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'parameter-default-type-mismatch')).toHaveLength(0);
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

	it('flags mismatched Exit statements after numeric line labels', () => {
		const src = 'Sub T()\n10 Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
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

describe('analyzeModule - procedure labels', () => {
	it('accepts forward and backward GoTo and GoSub labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'    GoTo done\n' +
			'start:\n' +
			'    GoSub done\n' +
			'    GoTo start\n' +
			'done:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('flags missing GoTo, GoSub, Resume, and On Error labels', () => {
		const src =
			'Sub T()\n' +
			'    GoTo MissingGoTo\n' +
			'    GoSub MissingGoSub\n' +
			'    Resume MissingResume\n' +
			'    On Error GoTo MissingHandler\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'MissingGoTo',
			'MissingGoSub',
			'MissingResume',
			'MissingHandler',
		]);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('keeps labels scoped to their enclosing procedure', () => {
		const src =
			'Sub A()\n' +
			'    GoTo Done\n' +
			'End Sub\n' +
			'Sub B()\n' +
			'Done:\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Done');
	});

	it('accepts explicit On Error forms that do not name a label', () => {
		const src =
			'Sub T()\n' +
			'    On Error Resume Next\n' +
			'    On Error GoTo 0\n' +
			'    On Error GoTo -1\n' +
			'    Resume\n' +
			'    Resume Next\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('accepts On Error disable forms after line numbers and inside single-line If statements', () => {
		const src =
			'Sub T(ByVal flag As Boolean)\n' +
			'10 On Error GoTo 0\n' +
			'20 If flag Then On Error GoTo 0 Else On Error GoTo Handler\n' +
			'30 On Error GoTo -1\n' +
			'Handler:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('validates On n GoTo and On n GoSub label lists', () => {
		const src =
			'Sub T(ByVal n As Long)\n' +
			'    On n GoTo First, MissingJump\n' +
			'    On n GoSub First, MissingSub\n' +
			'First:\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['MissingJump', 'MissingSub']);
	});

	it('validates labels on both sides of a single-line If Else statement', () => {
		const src = 'Sub T(ByVal flag As Boolean)\n    If flag Then GoTo MissingA Else GoTo MissingB\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['MissingA', 'MissingB']);
	});

	it('accepts numeric line labels and references', () => {
		const src =
			'Sub T()\n' +
			'    GoTo 10\n' +
			'10:\n' +
			'    On 1 GoTo 10\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('does not use labels from inactive conditional-compilation branches', () => {
		const src =
			'Sub T()\n' +
			'    GoTo MissingWhenInactive\n' +
			'#If VBA7 Then\n' +
			'MissingWhenInactive:\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: false } },
			}),
			'undefined-label',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MissingWhenInactive');
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

	it('accepts line-numbered Case statements inside Select Case', () => {
		const src =
			'Sub T()\n' +
			'10 Select Case x\n' +
			'20 Case 1, 2\n' +
			'30 x = 3\n' +
			'40 Case Else\n' +
			'50 x = 4\n' +
			'60 End Select\n' +
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

	it('uses the With receiver for unknown source-backed class members', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Delete\n' +
			'    End With\n' +
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
		expect(hits[0].message).toContain('Person.Delete');
	});

	it('uses the With receiver for source-backed class member assignment rules', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Age = 2\n' +
			'    End With\n' +
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
	});

	it('uses the With receiver for source-backed class member call rules', () => {
		const person = 'Public Sub Save(ByVal Count As Long)\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Save "bad"\n' +
			'        .Save()\n' +
			'    End With\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(1);
		expect(byCode(diagnostics, 'call-statement-forbids-parens')).toHaveLength(1);
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
		expect(spanText(src, hits[0])).toBe('Option');
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

