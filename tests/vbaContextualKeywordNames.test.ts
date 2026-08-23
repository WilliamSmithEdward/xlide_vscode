import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import { CONTEXTUAL_KEYWORDS } from '../src/analyzer/lexer/keywordTable';

// A CONTEXTUAL keyword is a word the VBE capitalizes in its statement context
// but MS-VBAL 3.3.5.2 does not reserve, so `Dim Text As String` and
// `Dim Error As Long` are both legal VBA. The reference scanner accepted only
// `identifier` tokens, so a name that spelled one of these was invisible: the
// ASSIGNMENT reported (`Text = 1` said Variable not defined) while the READ on
// the very next line said nothing. Same family as the keyword-named assignment
// target in #46.
//
// The other half of the change is the grammar guard. Each of these words has a
// position where it IS syntax - `On Error GoTo`, `Exit Property`, `For ... Step`,
// the `Open` mode and lock clauses - and scanning the word there would report a
// statement's own keyword as an undefined variable.

const NL = '\n';

function undeclared(body: readonly string[]): string[] {
	const source = ['Option Explicit', '', 'Public Sub P()', ...body, 'End Sub', ''].join(NL);
	return analyzeModule(source, { moduleName: 'M', knownIdentifiers: new Set<string>() })
		.filter((hit) => hit.code === 'undeclared-variable')
		.map((hit) => /'([^']+)'/.exec(hit.message)?.[1] ?? '?');
}

/**
 * The words this rule reads as names. `Property` opens a declaration and the
 * parser already refuses `Dim Property As Long`, so it can never reach a body
 * as a variable; `Error` names a runtime function, so a bare reference to it
 * resolves rather than reporting.
 */
const NAMEABLE = CONTEXTUAL_KEYWORDS.filter((w) => w !== 'Property' && w !== 'Error');

describe('a name that spells a contextual keyword', () => {
	it('reports an undeclared read, not only the assignment', () => {
		for (const word of NAMEABLE) {
			expect(undeclared([`    Debug.Print ${word}`]), `print ${word}`).toEqual([word]);
			expect(undeclared([`    MsgBox ${word}`]), `arg ${word}`).toEqual([word]);
			expect(
				undeclared(['    Dim n As Long', `    n = ${word} + 1`]),
				`expression ${word}`,
			).toEqual([word]);
		}
	});

	it('stays silent once the name is declared', () => {
		for (const word of CONTEXTUAL_KEYWORDS) {
			expect(
				undeclared([`    Dim ${word} As Long`, `    ${word} = 1`, `    Debug.Print ${word}`]),
				`declared ${word}`,
			).toEqual([]);
		}
	});

	it('reports the assignment target exactly once', () => {
		// The LHS skip is what stops the target being counted a second time as a
		// read. Broadening the scanner without broadening that skip double-reports.
		expect(undeclared(['    Text = 1'])).toEqual(['Text']);
	});

	it('leaves a member access alone', () => {
		expect(undeclared(['    Dim o As Object', '    Debug.Print o.Text'])).toEqual([]);
	});
});

describe('the grammar positions those words also occupy', () => {
	const legal: Record<string, readonly string[]> = {
		'For ... Step': ['    Dim i As Long', '    For i = 1 To 10 Step 2', '    Next i'],
		'On Error GoTo': ['    On Error GoTo H', '    Exit Sub', 'H:'],
		'On Error Resume Next': ['    On Error Resume Next'],
		'On Error GoTo 0': ['    On Error GoTo 0'],
		'Error statement': ['    Error 5'],
		// The statement is not always token 0 of the span the scanner sees.
		'Error after Then': ['    Dim b As Boolean', '    If b Then Error 5 Else Exit Sub'],
		'Error after a colon': ['    Dim x As Long', '    x = 1: Error 5'],
		'Open ... Lock Read Write': [
			'    Open "x" For Binary Access Read Write Lock Read Write As #1 Len = 1',
		],
		'Open ... Access Read Shared': ['    Open "x" For Input Access Read Shared As #1'],
		'Open ... For Output': ['    Open "x" For Output As #1'],
		'Open ... For Append': ['    Open "x" For Append As #1'],
		'Open ... For Random': ['    Open "x" For Random As #1'],
		'ReDim ... As Object': ['    Dim a() As Object', '    ReDim a(1) As Object'],
		'Dim ... As Object': ['    Dim o As Object', '    Set o = Nothing'],
		'TypeOf ... Is Object': ['    Dim o As Object', '    If TypeOf o Is Object Then', '    End If'],
		'Line Input': ['    Dim s As String', '    Line Input #1, s'],
	};

	for (const [label, body] of Object.entries(legal)) {
		it(`says nothing about ${label}`, () => {
			expect(undeclared(body)).toEqual([]);
		});
	}

	it('says nothing about Exit Property or the End Property footer', () => {
		const source = [
			'Option Explicit',
			'Public Property Get V() As Long',
			'    V = 1',
			'    Exit Property',
			'End Property',
			'',
		].join(NL);
		const hits = analyzeModule(source, {
			moduleName: 'C',
			moduleKind: 'class',
			knownIdentifiers: new Set<string>(),
		}).filter((hit) => hit.code === 'undeclared-variable');
		expect(hits).toEqual([]);
	});
});

describe('the Error function', () => {
	it('resolves bare and called, in both spellings', () => {
		// `Error[$]([errornumber])` reads the message text for a code and
		// defaults to the current Err.Number, so every one of these is a call.
		expect(undeclared(['    Dim s As String', '    s = Error(5)'])).toEqual([]);
		expect(undeclared(['    Dim s As String', '    s = Error$(5)'])).toEqual([]);
		expect(undeclared(['    Debug.Print Error'])).toEqual([]);
	});
});
