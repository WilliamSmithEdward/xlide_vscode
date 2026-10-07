import { describe, expect, it } from 'vitest';
import { analyzeModuleRulesIncremental } from '../src/analyzer/diagnostics/incrementalRules';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// Incremental rule re-analysis: body-only edits re-walk just the dirty
// procedure and must produce diagnostics identical to a full pass. Ordinary
// declarations/signatures invalidate consumers; directives, implicit member
// contracts, indirect declarations and external context retain full passes.

const BASE = [
	'Option Explicit',
	'Private mCount As Long',
	'',
	'Sub Alpha()',
	'    Dim a As Long',
	'    a = 1',
	'End Sub',
	'',
	'Sub Beta()',
	'    undeclaredBeta = 2',
	'End Sub',
	'',
	'Function Gamma() As Long',
	'    Gamma = mCount',
	'End Function',
].join('\n');

const FP = ['fp-a'] as const;

function run(source: string, state?: ReturnType<typeof analyzeVbaModuleSource>['rulesIncrementalState'], fingerprint: readonly unknown[] = FP) {
	return analyzeVbaModuleSource({
		source,
		moduleName: 'Module1',
		knownIdentifiers: new Set<string>(),
		rulesIncremental: { state, fingerprint },
	});
}

function key(r: ReturnType<typeof analyzeVbaModuleSource>): string {
	return r.diagnostics
		.map((d) => `${d.code}:${d.span.start}:${d.span.end}:${d.severity}:${d.message}`)
		.sort()
		.join('\n');
}

function full(source: string): ReturnType<typeof analyzeVbaModuleSource> {
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', knownIdentifiers: new Set<string>() });
}

describe('incremental rule re-analysis', () => {
	it('tracks header defaults that share a parameter name with a module constant', () => {
		const source = 'Private Const Limit As Long = 1\nSub F(Optional ByVal Limit As Long = Limit)\nMsgBox "abc"-\nEnd Sub\nSub P()\nF\nEnd Sub';
		const previous = run(source);
		const finding = previous.rulesIncrementalState!.walkDiagnosticsByProcedure[0].find(d => d.code === 'string-arithmetic-coercion');
		const changed = source.replace('Limit As Long = 1', 'Limit As Long = 0');
		const current = run(changed, previous.rulesIncrementalState);
		expect(key(current)).toBe(key(full(changed)));
		expect(current.rulesIncrementalState!.walkDiagnosticsByProcedure[0].find(d => d.code === finding!.code)).not.toBe(finding);
	});
	it.each(['Debug.Print 1', 'MsgBox "hello"'])('refreshes caller collection replay facts after an output edit (%s)', statement => {
		const source = 'Sub RemoveOne(ByVal p As Collection)\np.Remove 1\nEnd Sub\nSub Caller()\nDim c As New Collection\nc.Add 1\nRemoveOne c\nDebug.Print c(1)\nEnd Sub';
		const previous = run(source);
		expect(previous.diagnostics.some(d => d.code === 'collection-index-out-of-range')).toBe(true);
		const changed = source.replace('p.Remove 1', `p.Remove 1\n${statement}`);
		const current = run(changed, previous.rulesIncrementalState);
		expect(key(current)).toBe(key(full(changed)));
		expect(current.diagnostics.some(d => d.code === 'collection-index-out-of-range')).toBe(false);
		const restored = run(source, current.rulesIncrementalState);
		expect(key(restored)).toBe(key(full(source)));
		expect(restored.diagnostics.some(d => d.code === 'collection-index-out-of-range')).toBe(true);
	});
	it('matches full analysis across declaration edits and procedure additions, removals and reordering', () => {
		const source = 'Option Explicit\nPrivate m As Long\nPrivate Const Limit As Long = 1\n'
			+ 'Function F(ByRef n As Long) As Long\nF = Limit\nEnd Function\n'
			+ 'Sub Caller()\nDim n As Long\nm = 1\nDebug.Print 1 / m\nF n\nMissingAdded n\nEnd Sub\n'
			+ 'Sub Fixed()\nDim unused As Long\nMsgBox "abc"-\nEnd Sub\n'
			+ Array.from({ length: 8 }, (_, i) => `Sub Spare${i}()\nDebug.Print 1\nEnd Sub\n`).join('');
		const fixed = 'Sub Fixed()\nDim unused As Long\nMsgBox "abc"-\nEnd Sub\n';
		for (const changed of [
			source.replace('Private m As Long', 'Private m As String'),
			source.replace('Limit As Long = 1', 'Limit As Long = 0'),
			source.replace('F(ByRef n As Long)', 'F(ByRef n As Long, ByVal extra As Long)'),
			source.replace('Sub Fixed()', 'Sub MissingAdded(ByRef n As Long)\nn = 2\nEnd Sub\nSub Fixed()'),
			source.replace('Function F(ByRef n As Long) As Long\nF = Limit\nEnd Function\n', ''),
			source.replace(fixed, '') + fixed,
			source.replace('Private m As Long', 'Private m As Long\nPrivate MissingAdded As Long'),
		]) {
			const previous = run(source);
			const current = run(changed, previous.rulesIncrementalState);
			expect(current.rulesIncrementalMode).toBe('incremental');
			expect(key(current)).toBe(key(full(changed)));
			const restored = run(source, current.rulesIncrementalState);
			expect(restored.rulesIncrementalMode).toBe('incremental');
			expect(key(restored)).toBe(key(full(source)));
		}
	});
	it.each([
		'Private Const Base As Long = 1\nPrivate Const Derived As Long = Base\nSub P()\nDebug.Print Derived\nEnd Sub',
		'Private Const Base As Long = 1\nType T\nv(Base) As Long\nEnd Type\nSub P()\nDim v As T\nEnd Sub',
	])('retains full analysis for indirect declaration dependencies', source => {
		const previous = run(source);
		const changed = source.replace('Base As Long = 1', 'Base As Long = 0');
		const current = run(changed, previous.rulesIncrementalState);
		expect(current.rulesIncrementalMode).toBe('full');
		expect(key(current)).toBe(key(full(changed)));
	});
	it.each(['', "    ' comment only\n", '    Debug.Print 1\n'])('keeps first/last/empty body edits incremental (%j)', body => {
		const source = `Option Explicit\nSub A()\n${body}End Sub\nSub B()\nMsgBox "abc"-\nEnd Sub`;
		let previous = run(source);
		for (const changed of [source.replace('Sub A()\n', 'Sub A()\n    missingFirst = 1\n'),
			source.replace('End Sub\nSub B', '    missingLast = 1\nEnd Sub\nSub B'), source]) {
			const current = run(changed, previous.rulesIncrementalState);
			expect(current.rulesIncrementalMode).toBe('incremental');
			expect(key(current)).toBe(key(full(changed)));
			previous = current;
		}
	});
	it('invalidates changed exported attributes', () => {
		const source = 'Public Function Value() As Long\nAttribute Value.VB_UserMemId = 0\nValue = 1\nEnd Function';
		const previous = run(source);
		for (const changed of [source.replace('= 0', '= 1'), source.replace('Value() As Long', 'Value(ByVal n As Long) As Long'), 'Sub P()\nEnd Sub']) {
			const current = run(changed, previous.rulesIncrementalState);
			expect(current.rulesIncrementalMode).toBe('full');
			expect(key(current)).toBe(key(full(changed)));
		}
	});
	it('invalidates implicit default-member consumers when a setter signature changes', () => {
		const source = 'Property Get Item() As Long\nAttribute Item.VB_UserMemId = 0\nItem = 1\nEnd Property\nProperty Let Item(ByVal v As Long)\nEnd Property\nSub P()\nDim c As Module1\nDebug.Print c\nEnd Sub';
		const previous = run(source);
		for (const changed of [source.replace('Property Let Item(ByVal v As Long)', 'Property Let Item(ByVal v As String)'),
			source.replace('Property Let Item(ByVal v As Long)\nEnd Property\n', '')]) {
			const current = run(changed, previous.rulesIncrementalState);
			expect(current.rulesIncrementalMode).toBe('full');
			expect(key(current)).toBe(key(full(changed)));
		}
	});
	it.each(['End Sub', ''])('publishes and clears changed/missing procedure closers incrementally (%j)', closer => {
		const source = 'Function F() As Long\nF = 1\nEnd Function\nSub P()\nMsgBox "abc"-\nEnd Sub';
		const previous = run(source);
		const changed = source.replace('End Function', closer);
		const current = run(changed, previous.rulesIncrementalState);
		expect(current.rulesIncrementalMode).toBe('incremental');
		expect(key(current)).toBe(key(full(changed)));
		const restored = run(source, current.rulesIncrementalState);
		expect(restored.rulesIncrementalMode).toBe('incremental');
		expect(key(restored)).toBe(key(full(source)));
	});
	it('resolves caller dependencies through local shadows and typed receivers', () => {
		const source = 'Function Value() As Long\nValue = 1\nEnd Function\n'
			+ 'Sub Direct()\nDebug.Print 10 / Value()\nEnd Sub\n'
			+ 'Sub Shadow(ByVal Value As Long)\nDebug.Print Value\nMsgBox "abc"-\nEnd Sub\n'
			+ 'Sub External(ByVal wb As Workbook)\nDebug.Print wb.Value\nMsgBox "abc"-\nEnd Sub\n'
			+ 'Sub Own(ByVal receiver As C)\nDebug.Print receiver.Value()\nMsgBox "abc"-\nEnd Sub\n'
			+ 'Sub Dynamic(ByVal receiver As Object)\nDebug.Print receiver.Value()\nMsgBox "abc"-\nEnd Sub\n'
			+ 'Sub Named()\nApplication.Run "Value"\nMsgBox "abc"-\nEnd Sub\n'
			+ Array.from({ length: 5 }, (_, i) => `Sub Spare${i}()\nDebug.Print 1\nEnd Sub\n`).join('');
		const opts = { moduleName: 'C', moduleKind: 'class' as const };
		const previous = analyzeModuleRulesIncremental(source, opts, undefined, FP);
		const changed = source.replace('Value = 1', 'Value = 0');
		const current = analyzeModuleRulesIncremental(changed, opts, previous.state, FP);
		expect(current.mode).toBe('incremental');
		const fullResult = analyzeModuleRulesIncremental(changed, opts, undefined, FP);
		const keys = (diagnostics: typeof current.diagnostics) => diagnostics.map(d => `${d.code}:${d.span.start}:${d.span.end}:${d.severity}:${d.message}`).sort();
		expect(keys(current.diagnostics)).toEqual(keys(fullResult.diagnostics));
		expect(current.diagnostics.some(d => d.code === 'division-by-zero')).toBe(true);
		for (const [procedure, reused] of [['Shadow', true], ['External', true], ['Own', false], ['Dynamic', false], ['Named', false]] as const) {
			const offset = source.indexOf('MsgBox "abc"-', source.indexOf(`Sub ${procedure}(`));
			const before = previous.diagnostics.find(d => d.code === 'string-arithmetic-coercion' && d.span.start >= offset && d.span.start < offset + 20);
			expect(before).toBeDefined();
			const after = current.diagnostics.find(d => d.code === before!.code && d.span.start === before!.span.start);
			if (reused) { expect(after).toBe(before); } else { expect(after).not.toBe(before); }
		}
	});
	it('refreshes information findings and module uses while retaining untouched local findings', () => {
		const source = 'Option Explicit\nPrivate mUnused As Long\nSub A()\nDebug.Print 0\nDebug.Print 1\nEnd Sub\nSub B()\nDim untouched As Long\nEnd Sub';
		const initial = run(source);
		const changed = source.replace('Debug.Print 1', 'Dim written As Long\nwritten = 1\nDebug.Print mUnused\nExit Sub\nDebug.Print 2');
		const edited = run(changed, initial.rulesIncrementalState);
		expect(edited.rulesIncrementalMode).toBe('incremental');
		expect(key(edited)).toBe(key(full(changed)));
		expect(edited.diagnostics.some(d => d.code === 'unreachable-code')).toBe(true);
		expect(edited.diagnostics.some(d => d.code === 'variable-never-read')).toBe(true);
		expect(edited.diagnostics.some(d => d.code === 'unused-variable' && d.message.includes('mUnused'))).toBe(false);
		expect(edited.diagnostics.some(d => d.code === 'unused-variable' && d.message.includes('untouched'))).toBe(true);
		const cleared = run(source, edited.rulesIncrementalState);
		expect(key(cleared)).toBe(key(full(source)));
	});
	it('re-analyzes only a changed body and matches the full pass exactly', () => {
		const base = run(BASE);
		expect(base.rulesIncrementalMode).toBe('full');
		// Introduce a new error inside Alpha's body.
		const edited = BASE.replace('    a = 1', '    a = 1\n    undeclaredAlpha = 3');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('incremental');
		expect(key(incr)).toBe(key(full(edited)));
		// The pre-existing error in the untouched Beta must survive the splice.
		expect(incr.diagnostics.some((d) => d.message.includes('undeclaredBeta'))).toBe(true);
		expect(incr.diagnostics.some((d) => d.message.includes('undeclaredAlpha'))).toBe(true);
	});

	it('clears a fixed error in the edited body and keeps others', () => {
		const base = run(BASE);
		const edited = BASE.replace('    undeclaredBeta = 2', '    mCount = 2');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('full');
		expect(key(incr)).toBe(key(full(edited)));
		expect(incr.diagnostics.some((d) => d.message.includes('undeclaredBeta'))).toBe(false);
	});

	it('chains state across successive edits', () => {
		let source = BASE;
		let r = run(source);
		for (let i = 0; i < 3; i += 1) {
			source = source.replace('    a = 1', `    a = 1\n    a = a + ${i}`);
			r = run(source, r.rulesIncrementalState);
			expect(r.rulesIncrementalMode).toBe('incremental');
			expect(key(r)).toBe(key(full(source)));
		}
	});

	it('rechecks a changed signature and its consumers incrementally', () => {
		const base = run(BASE);
		const edited = BASE.replace('Sub Alpha()', 'Sub Alpha(ByVal n As Long)');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('incremental');
		expect(key(incr)).toBe(key(full(edited)));
	});

	it('keeps an unused declaration addition incremental', () => {
		const base = run(BASE);
		const edited = BASE.replace('Private mCount As Long', 'Private mCount As Long\nPrivate mOther As String');
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('incremental');
		expect(key(incr)).toBe(key(full(edited)));
	});

	it('keeps a procedure addition incremental', () => {
		const base = run(BASE);
		const edited = `${BASE}\n\nSub Delta()\nEnd Sub`;
		const incr = run(edited, base.rulesIncrementalState);
		expect(incr.rulesIncrementalMode).toBe('incremental');
		expect(key(incr)).toBe(key(full(edited)));
	});

	it('falls back to full on a fingerprint change (cross-module inputs moved)', () => {
		const base = run(BASE);
		const edited = BASE.replace('    a = 1', '    a = 2');
		const incr = run(edited, base.rulesIncrementalState, ['fp-b']);
		expect(incr.rulesIncrementalMode).toBe('full');
		expect(key(incr)).toBe(key(full(edited)));
	});

    it('rechecks direct return-value consumers without invalidating unrelated transitive findings', () => {
        const source = 'Function F() As Long\nF = 1\nEnd Function\nFunction G() As Long\nG = F()\nEnd Function\nSub P()\nDebug.Print G()\nMsgBox "abc"-\nEnd Sub';
        const previous = analyzeModuleRulesIncremental(source, {}, undefined, FP);
        const edited = source.replace('F = 1', 'F = 0');
        const incremental = analyzeModuleRulesIncremental(edited, {}, previous.state, FP);
        const oldFinding = previous.diagnostics.find(d => d.code === 'string-arithmetic-coercion');
        expect(oldFinding).toBeDefined();
        // G's exported facts do not recursively evaluate F's return value.
        // P therefore keeps its unrelated finding at identical offsets.
        expect(incremental.diagnostics.find(d => d.code === 'string-arithmetic-coercion')).toBe(oldFinding);
        expect(incremental.diagnostics.map(d => `${d.code}:${d.span.start}:${d.message}`).sort())
            .toEqual(analyzeModuleRulesIncremental(edited, {}, undefined, FP).diagnostics.map(d => `${d.code}:${d.span.start}:${d.message}`).sort());
    });

	it.each(['DoEvents', 'MsgBox "hello"'])('propagates changes to transitive module-variable effects (%s)', statement => {
		const source = 'Private m As Long\nSub F()\nDebug.Print 1\nEnd Sub\nSub G()\nF\nEnd Sub\nSub P()\nm = 0\nG\nDebug.Print 1 / m\nEnd Sub';
		const previous = run(source);
		expect(previous.diagnostics.some(d => d.code === 'division-by-zero')).toBe(true);
		const changed = source.replace('Debug.Print 1\nEnd Sub', statement + '\nEnd Sub');
		const current = run(changed, previous.rulesIncrementalState);
		expect(current.rulesIncrementalMode).toBe('incremental');
		expect(key(current)).toBe(key(full(changed)));
		expect(current.diagnostics.some(d => d.code === 'division-by-zero')).toBe(false);
		const cleared = run(source, current.rulesIncrementalState);
		expect(key(cleared)).toBe(key(full(source)));
		expect(cleared.diagnostics.some(d => d.code === 'division-by-zero')).toBe(true);
	});

	it('propagates changes to transitive argument-preservation effects', () => {
		const source = 'Sub F(ByRef n As Long)\nDebug.Print 1\nEnd Sub\nFunction G() As Long\nDim n As Long\nF n\nG = n\nEnd Function\nSub P()\nDebug.Print 1 / G()\nEnd Sub';
		const previous = run(source);
		expect(previous.diagnostics.some(d => d.code === 'division-by-zero')).toBe(true);
		const changed = source.replace('Debug.Print 1\nEnd Sub', 'n = 1\nEnd Sub');
		const current = run(changed, previous.rulesIncrementalState);
		expect(current.rulesIncrementalMode).toBe('incremental');
		expect(key(current)).toBe(key(full(changed)));
		expect(current.diagnostics.some(d => d.code === 'division-by-zero')).toBe(false);
	});

    it('splices and shifts eager procedure findings without duplicating them', () => {
        const source = 'Sub A()\nDebug.Print 1\nEnd Sub\nSub B()\nDim s As String * 0\nDim a(1) As Long\nDebug.Print a(2)\nDim c As Collection\nDebug.Print c.Count\nEnd Sub';
        const previous = run(source);
        const edited = source.replace('Debug.Print 1', 'Debug.Print 12345\nDebug.Print 2');
        const incremental = run(edited, previous.rulesIncrementalState);
        expect(incremental.rulesIncrementalMode).toBe('incremental');
        expect(key(incremental)).toBe(key(full(edited)));
        for (const code of ['fixed-length-string-size', 'array-subscript-out-of-bounds', 'object-variable-not-set']) {
            expect(incremental.diagnostics.filter(d => d.code === code)).toHaveLength(1);
        }
    });

    it('invalidates cached runtime findings when a module variable becomes written', () => {
        const source = 'Private m As Collection\nSub A()\nDebug.Print 1\nEnd Sub\nSub B()\nDebug.Print m.Count\nEnd Sub';
        const previous = run(source);
        expect(previous.diagnostics.some(d => d.code === 'object-variable-not-set')).toBe(true);
        const edited = source.replace('Debug.Print 1', 'Set m = New Collection');
        const incremental = run(edited, previous.rulesIncrementalState);
        expect(incremental.rulesIncrementalMode).toBe('full');
        expect(key(incremental)).toBe(key(full(edited)));
        expect(incremental.diagnostics.some(d => d.code === 'object-variable-not-set')).toBe(false);
    });

    it('invalidates when a directive inside a body changes', () => {
        const source = 'Option Explicit\nSub A()\n#Const FLAG = True\nDebug.Print 1\nEnd Sub\nSub B()\n#If FLAG Then\nunknown = 1\n#End If\nEnd Sub';
        const previous = run(source);
        const edited = source.replace('FLAG = True', 'FLAG = False');
        const incremental = run(edited, previous.rulesIncrementalState);
        expect(incremental.rulesIncrementalMode).toBe('full');
        expect(key(incremental)).toBe(key(full(edited)));
    });

    it('separates error-only and full state even with the same caller fingerprint', () => {
        const source = 'Sub P()\nDim unused As Long\nDebug.Print 1\nEnd Sub';
        const previous = analyzeModuleRulesIncremental(source, { errorsOnly: true }, undefined, FP);
        const complete = analyzeModuleRulesIncremental(source, {}, previous.state, FP);
        expect(complete.mode).toBe('full');
        expect(complete.diagnostics.some(d => d.code === 'unused-variable')).toBe(true);
    });

    it('keeps shadowed output calls in the callee dependency fingerprint', () => {
        const source = 'Sub MsgBox(ByVal s As String)\nDebug.Print s\nEnd Sub\nFunction F() As Long\nF = 1\nMsgBox "old"\nEnd Function\nSub P()\nDebug.Print F()\nMsgBox "abc"-\nEnd Sub';
        const previous = analyzeModuleRulesIncremental(source, {}, undefined, FP);
        const edited = source.replace('MsgBox "old"', 'MsgBox "new"');
        const incremental = analyzeModuleRulesIncremental(edited, {}, previous.state, FP);
        const finding = previous.diagnostics.find(d => d.code === 'string-arithmetic-coercion');
        expect(finding).toBeDefined();
        expect(incremental.diagnostics.find(d => d.code === finding!.code && d.span.start === finding!.span.start)).not.toBe(finding);
    });

    it('rechecks string-based callers when their target body changes', () => {
        const source = 'Sub Target()\nDim n As Long\nn = 1\nEnd Sub\nSub P()\nCallByName Me, "Target", VbMethod\nMsgBox "abc"-\nEnd Sub';
        const opts = { moduleKind: 'class' as const, moduleName: 'C' };
        const previous = analyzeModuleRulesIncremental(source, opts, undefined, FP);
        const edited = source.replace('n = 1', 'n = 2');
        const incremental = analyzeModuleRulesIncremental(edited, opts, previous.state, FP);
        const finding = previous.diagnostics.find(d => d.code === 'string-arithmetic-coercion');
        expect(finding).toBeDefined();
        expect(incremental.diagnostics.find(d => d.code === finding!.code)).not.toBe(finding);
    });

});
