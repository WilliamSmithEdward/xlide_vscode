import { describe, it, expect } from 'vitest';
import {
	bareCallStatementTarget,
	callableCompletionShouldInsertParens,
	explicitCallStatementArgumentWithoutParens,
	explicitCallStatementTarget,
	findActiveCallSite,
	standaloneEmptyParenthesizedCallStatement,
} from '../src/analyzer';

function at(src: string): { source: string; offset: number } {
	const offset = src.indexOf('|');
	if (offset < 0) {
		throw new Error('missing | marker');
	}
	return { source: src.slice(0, offset) + src.slice(offset + 1), offset };
}

function lineSpan(source: string, text: string): { start: number; end: number } {
	const start = source.indexOf(text);
	if (start < 0) {
		throw new Error(`missing ${text}`);
	}
	return { start, end: start + text.length };
}

describe('VBA call context', () => {
	it('classifies parenless member call statements for signature help', () => {
		const { source, offset } = at('Sub T()\nWorkbooks.Open "a.xlsx", |\nEnd Sub');
		const site = findActiveCallSite(source, offset);

		expect(site).toMatchObject({
			calleeName: 'Open',
			isMember: true,
			isExplicitCall: false,
			activeParameter: 1,
		});
	});

	it('classifies leading-dot parenless member calls inside With blocks', () => {
		const { source, offset } = at('Sub T()\nWith Range("A1")\n    .Offset |\nEnd With\nEnd Sub');
		const site = findActiveCallSite(source, offset);

		expect(site).toMatchObject({
			calleeName: 'Offset',
			isMember: true,
			isExplicitCall: false,
			activeParameter: 0,
		});
	});

	it('does not confuse statement keywords with parenless call targets', () => {
		const { source, offset } = at('Sub T()\nOpen "a.txt" For Input As #1, |\nEnd Sub');
		expect(findActiveCallSite(source, offset)).toBeUndefined();
	});

	it('detects explicit Call targets for parenthesized calls', () => {
		const { source, offset } = at('Sub T()\nCall SaveFile(|\nEnd Sub');
		const site = findActiveCallSite(source, offset);

		expect(site).toMatchObject({
			calleeName: 'SaveFile',
			isMember: false,
			isExplicitCall: true,
			activeParameter: 0,
		});
	});

	it('uses one paren-insertion rule for bare and member completions', () => {
		const statementCall = 'Sub T()\n    mySub\nEnd Sub\n';
		const explicitCall = 'Sub T()\n    Call mySub\nEnd Sub\n';
		const expressionCall = 'Sub T()\n    value = Left\nEnd Sub\n';
		const memberStatement = 'Sub T()\n    Application.Calculate\nEnd Sub\n';

		expect(callableCompletionShouldInsertParens(statementCall, statementCall.indexOf('mySub') + 5))
			.toBe(false);
		expect(callableCompletionShouldInsertParens(explicitCall, explicitCall.indexOf('mySub') + 5))
			.toBe(true);
		expect(callableCompletionShouldInsertParens(expressionCall, expressionCall.indexOf('Left') + 4))
			.toBe(true);
		expect(callableCompletionShouldInsertParens(
			memberStatement,
			memberStatement.indexOf('Calculate') + 9,
		)).toBe(false);
	});

	it('extracts bare call-statement targets for diagnostics', () => {
		const source = 'Sub T()\n    MsgBox "hi"\n    Call SaveFile(1)\nEnd Sub\n';

		expect(bareCallStatementTarget(source, lineSpan(source, 'MsgBox "hi"')))
			.toMatchObject({ name: 'MsgBox' });
		expect(bareCallStatementTarget(source, lineSpan(source, 'Call SaveFile(1)')))
			.toMatchObject({ name: 'SaveFile' });
	});

	it('keeps labels, assignments, members, and implicit Application forms out of bare diagnostics', () => {
		const source =
			'Sub T()\n' +
			'done:\n' +
			'    value = Foo\n' +
			'    Application.Calculate\n' +
			'    Range("A1")\n' +
			'End Sub\n';

		expect(bareCallStatementTarget(source, lineSpan(source, 'done'))).toBeUndefined();
		expect(bareCallStatementTarget(source, lineSpan(source, 'value = Foo'))).toBeUndefined();
		expect(bareCallStatementTarget(source, lineSpan(source, 'Application.Calculate')))
			.toBeUndefined();
		expect(bareCallStatementTarget(source, lineSpan(source, 'Range("A1")'))).toBeUndefined();
	});

	it('extracts explicit Call targets and unparenthesized Call arguments for diagnostics', () => {
		const source =
			'Sub T()\n' +
			'    Call DoEvents\n' +
			'    Call Application.Calculate 1\n' +
			'    Call Workbooks(1).Sheets(1).Move before:=Sheets(2)\n' +
			'End Sub\n';

		expect(explicitCallStatementTarget(source, lineSpan(source, 'Call DoEvents')))
			.toMatchObject({ name: 'DoEvents' });
		expect(
			explicitCallStatementArgumentWithoutParens(
				source,
				lineSpan(source, 'Call Application.Calculate 1'),
			),
		).toEqual(lineSpan(source, '1'));
		expect(
			explicitCallStatementArgumentWithoutParens(
				source,
				lineSpan(source, 'Call Workbooks(1).Sheets(1).Move before:=Sheets(2)'),
			),
		).toEqual(lineSpan(source, 'before'));
	});

	it('classifies complete standalone empty-parentheses call statements', () => {
		const source =
			'Sub T()\n' +
			'    myFunction()\n' +
			'    Application.Calculate()\n' +
			'    Workbooks(1).Sheets(1).Range()\n' +
			'    value = myFunction()\n' +
			'End Sub\n';

		expect(
			standaloneEmptyParenthesizedCallStatement(
				source,
				lineSpan(source, 'myFunction()'),
			),
		).toMatchObject({ name: 'myFunction', isMember: false });
		expect(
			standaloneEmptyParenthesizedCallStatement(
				source,
				lineSpan(source, 'Application.Calculate()'),
			),
		).toMatchObject({ name: 'Calculate', isMember: true, startsWithLeadingDot: false });
		expect(
			standaloneEmptyParenthesizedCallStatement(
				source,
				lineSpan(source, 'Workbooks(1).Sheets(1).Range()'),
			),
		).toMatchObject({ name: 'Range', isMember: true });
		expect(
			standaloneEmptyParenthesizedCallStatement(
				source,
				lineSpan(source, 'value = myFunction()'),
			),
		).toBeUndefined();
	});

	it('classifies leading-dot empty-parentheses call statements inside With blocks', () => {
		const source = 'Sub T()\n    With ActiveSheet\n        .Calculate()\n    End With\nEnd Sub\n';

		expect(
			standaloneEmptyParenthesizedCallStatement(
				source,
				lineSpan(source, '.Calculate()'),
			),
		).toMatchObject({
			name: 'Calculate',
			isMember: true,
			startsWithLeadingDot: true,
		});
	});
});
