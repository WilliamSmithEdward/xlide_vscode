import { describe, it, expect } from 'vitest';
import {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
	runtimeAllowsExplicitCall,
	VBA_RUNTIME_CONSTANTS,
	VBA_RUNTIME_FUNCTIONS,
	VBA_RUNTIME_OBJECTS,
	resolveIdentifierCompletions,
} from '../src/analyzer';

describe('VBA runtime metadata', () => {
	it('resolves built-ins case-insensitively', () => {
		expect(resolveRuntimeFunction('MsgBox')?.name).toBe('MsgBox');
		expect(resolveRuntimeFunction('msgbox')?.name).toBe('MsgBox');
		expect(resolveRuntimeFunction('LEFT')?.returns).toBe('String');
		expect(resolveRuntimeFunction('CLng')?.returns).toBe('Long');
		expect(resolveRuntimeFunction('NotARealFunction')).toBeUndefined();
		expect(resolveRuntimeObject('Err')?.type).toBe('VBA.ErrObject');
		expect(resolveRuntimeObject('ERR')?.name).toBe('Err');
		expect(resolveRuntimeObject('NotARealObject')).toBeUndefined();
	});

	it('every entry has a signature and is marked verified', () => {
		for (const f of VBA_RUNTIME_FUNCTIONS) {
			expect(f.name.length).toBeGreaterThan(0);
			expect(f.signature.length).toBeGreaterThan(0);
			expect(f.source).toBe('verified');
			if (f.kind === 'function') {
				expect(f.signature).toContain('As ');
			}
		}
	});

	it('excludes intrinsic data-type names to avoid type/function ambiguity', () => {
		for (const name of ['String', 'Date', 'Time', 'Error']) {
			expect(resolveRuntimeFunction(name)).toBeUndefined();
		}
	});

	it('includes verified gap-filled built-ins from the VBA library', () => {
		expect(resolveRuntimeFunction('StrReverse')?.returns).toBe('String');
		expect(resolveRuntimeFunction('FormatCurrency')?.kind).toBe('function');
		expect(resolveRuntimeFunction('Dir')?.returns).toBe('String');
		expect(resolveRuntimeFunction('FreeFile')?.returns).toBe('Integer');
		expect(resolveRuntimeFunction('Pmt')?.returns).toBe('Double');
		expect(resolveRuntimeFunction('CallByName')?.name).toBe('CallByName');
		// File-system commands are statements, not value-returning functions.
		expect(resolveRuntimeFunction('Kill')?.kind).toBe('statement');
		expect(resolveRuntimeFunction('MkDir')?.kind).toBe('statement');
	});

	it('records runtime functions that cannot be explicit Call targets', () => {
		const doEvents = resolveRuntimeFunction('DoEvents');
		const msgBox = resolveRuntimeFunction('MsgBox');

		expect(doEvents).toBeDefined();
		expect(msgBox).toBeDefined();
		expect(runtimeAllowsExplicitCall(doEvents!)).toBe(false);
		expect(runtimeAllowsExplicitCall(msgBox!)).toBe(true);
	});

	it('records Array ArgList as a zero-or-more ParamArray', () => {
		const array = resolveRuntimeFunction('Array');

		expect(array?.params).toEqual([
			{
				name: 'ArgList',
				type: 'Variant',
				optional: true,
				paramArray: true,
			},
		]);
	});

	it('resolves built-in constants case-insensitively', () => {
		expect(resolveRuntimeConstant('vbOKOnly')?.type).toBe('VbMsgBoxStyle');
		expect(resolveRuntimeConstant('VBOKONLY')?.name).toBe('vbOKOnly');
		expect(resolveRuntimeConstant('vbCrLf')?.type).toBe('String');
		expect(resolveRuntimeConstant('notAConstant')).toBeUndefined();
	});

	it('every constant entry has a name and is marked verified', () => {
		for (const constant of VBA_RUNTIME_CONSTANTS) {
			expect(constant.name.length).toBeGreaterThan(0);
			expect(constant.source).toBe('verified');
		}
	});

	it('models intrinsic runtime objects as exhaustive member surfaces', () => {
		for (const object of VBA_RUNTIME_OBJECTS) {
			expect(object.name.length).toBeGreaterThan(0);
			expect(object.type).toMatch(/^VBA\./);
			expect(object.source).toBe('verified');
			expect(object.exhaustive).toBe(true);
			expect(object.members.length).toBeGreaterThan(0);
		}
		const err = resolveRuntimeObject('Err');
		expect(err?.members.map((member) => member.name)).toContain('Raise');
		expect(err?.members.find((member) => member.name === 'Raise')?.signature).toContain(
			'Number As Long',
		);
	});

});

describe('identifier completion - runtime built-ins', () => {
	const src = 'Sub T()\n    \nEnd Sub\n';
	const offset = src.indexOf('    \n') + 4;

	it('offers built-in functions at a statement start', () => {
		const names = resolveIdentifierCompletions(src, offset, {
			moduleName: 'M',
		}).map((c) => c.name);
		expect(names).toContain('MsgBox');
		expect(names).toContain('Array');
		expect(names).toContain('RGB');
		expect(names).toContain('Err');
	});

	it('filters built-ins by the typed prefix', () => {
		const typed = 'Sub T()\n    Msg\nEnd Sub\n';
		const off = typed.indexOf('Msg') + 3;
		const names = resolveIdentifierCompletions(typed, off, {
			moduleName: 'M',
		}).map((c) => c.name);
		expect(names).toContain('MsgBox');
		expect(names).not.toContain('Array');
	});

	it('does not offer runtime functions that VBE rejects as explicit Call targets', () => {
		const typed = 'Sub T()\n    Call DoE\nEnd Sub\n';
		const off = typed.indexOf('DoE') + 3;
		const names = resolveIdentifierCompletions(typed, off, {
			moduleName: 'M',
		}).map((c) => c.name);
		expect(names).not.toContain('DoEvents');

		const expression = 'Sub T()\n    value = DoE\nEnd Sub\n';
		const expressionOff = expression.indexOf('DoE') + 3;
		const expressionNames = resolveIdentifierCompletions(expression, expressionOff, {
			moduleName: 'M',
		}).map((c) => c.name);
		expect(expressionNames).toContain('DoEvents');
	});

	it('shows verified signatures and parameter metadata for runtime completions', () => {
		const typed = 'Sub T()\n    Le\nEnd Sub\n';
		const off = typed.indexOf('Le') + 2;
		const got = resolveIdentifierCompletions(typed, off, {
			moduleName: 'M',
		});
		const left = got.find((item) => item.name === 'Left');
		expect(left?.detail).toBe('Left(String, Length) As String');
		expect(left?.documentation).toContain('**VBA runtime function**');
		expect(left?.documentation).toContain('```vba\nLeft(String, Length) As String\n```');
		expect(left?.documentation).toContain('`String` As `String`');
		expect(left?.documentation).toContain('`Length` As `Long`');

		const errSrc = 'Sub T()\n    Er\nEnd Sub\n';
		const errItems = resolveIdentifierCompletions(errSrc, errSrc.indexOf('Er') + 2, {
			moduleName: 'M',
		});
		const err = errItems.find((item) => item.name === 'Err');
		expect(err?.detail).toBe('VBA runtime object As ErrObject');
		expect(err?.documentation).toContain('Err As ErrObject');
	});

	it('offers runtime and host constants once a constant-like prefix is typed', () => {
		const vb = 'Sub T()\n    vb\nEnd Sub\n';
		const vbItems = resolveIdentifierCompletions(vb, vb.indexOf('vb') + 2, {
			moduleName: 'M',
		});
		const vbOkOnly = vbItems.find((item) => item.name === 'vbOKOnly');
		expect(vbOkOnly?.kind).toBe('constant');
		expect(vbOkOnly?.detail).toBe('VBA constant As VbMsgBoxStyle');
		expect(vbOkOnly?.documentation).toContain('Const vbOKOnly As VbMsgBoxStyle = 0');

		const xl = 'Sub T()\n    xl\nEnd Sub\n';
		const xlItems = resolveIdentifierCompletions(xl, xl.indexOf('xl') + 2, {
			moduleName: 'M',
		});
		const xlUp = xlItems.find((item) => item.name === 'xlUp');
		expect(xlUp?.kind).toBe('constant');
		expect(xlUp?.detail).toBe('Excel/Office constant As XlDirection');
		expect(xlUp?.documentation).toContain('Const xlUp As XlDirection = -4162');

		const mso = 'Sub T()\n    msoLine\nEnd Sub\n';
		const msoItems = resolveIdentifierCompletions(mso, mso.indexOf('msoLine') + 7, {
			moduleName: 'M',
		});
		const msoLineDash = msoItems.find((item) => item.name === 'msoLineDash');
		expect(msoLineDash?.kind).toBe('constant');
		expect(msoLineDash?.detail).toBe('Excel/Office constant As MsoLineDashStyle');
		expect(msoLineDash?.documentation).toContain('Const msoLineDash As MsoLineDashStyle = 4');
	});

	it('can be disabled via includeRuntime', () => {
		const names = resolveIdentifierCompletions(src, offset, {
			moduleName: 'M',
			includeRuntime: false,
		}).map((c) => c.name);
		expect(names).not.toContain('MsgBox');
	});
});
