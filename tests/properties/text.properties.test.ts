import { describe, it } from 'vitest';
import fc from 'fast-check';
import { formatVbaModule } from '../../src/analyzer/format/formatModule';
import { tokenize } from '../../src/analyzer/lexer/tokenize';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { FrmHeaderError, parseFrmHeader, printFrmHeader } from '../../src/vba/vb6/frmHeader';
import { VbpManifestError, parseVbpManifest, printVbpManifest } from '../../src/vba/vb6/vbpProject';
import { propertySettings } from './settings';

// Properties of the readers of VBA and VB6 text: the lexer, the parser, the
// formatter, and the VB6 form header and project manifest. Module text is
// whatever a workbook or a project folder carries, so each reader must keep
// its promise for text nobody wrote a case for. The Fuzz workflow raises
// XLIDE_PROPERTY_RUNS; see ./settings.ts.

/** Fragments that steer generated lines into the lexer's and parser's decisions. */
const FRAGMENTS = [
	'Sub', 'Function', 'Property Get', 'Property Let', 'Property Set', 'End Sub', 'End Function',
	'End Property', 'Public', 'Private', 'Friend', 'Static', 'Dim', 'ReDim Preserve', 'Const', 'As',
	'Long', 'String * 10', 'Variant', 'New', 'If', 'Then', 'Else', 'ElseIf', 'End If', 'EndIf',
	'For', 'Each', 'In', 'To', 'Step', 'Next', 'Do', 'Loop', 'While', 'Until', 'Wend', 'With',
	'End With', 'Select Case', 'Case Is >', 'Case Else', 'End Select', 'Type', 'End Type', 'Enum',
	'End Enum', 'On Error GoTo', 'Resume Next', 'GoSub', 'Exit Sub', 'Declare PtrSafe', 'Lib',
	'Option Explicit', 'Attribute VB_Name =', 'Rem', "'", '"', '""', '"a""b"', ' _', '_', ':',
	'.', '!', '#If', '#Else', '#End If', '#Const', '#1/2/2003#', '&H7FFF&', '&O17', '1.5E+10',
	'1#', 'x%', 'y$', '[odd name]', 'step', 'error', 'label1:', 'étape', 'Привет', 'x', '=',
	'(', ')', ',', ';', '&', '+', '-', '*', '/', '\\', '^', 'Mod', 'And', 'Not', 'Like', 'Is',
	'Nothing', 'Me', ' ', '\t', ' ', '　',
];

const vbaLine = fc
	.array(fc.oneof(
		{ weight: 5, arbitrary: fc.constantFrom(...FRAGMENTS) },
		{ weight: 1, arbitrary: fc.string({ unit: 'binary', maxLength: 6 }) },
	), { maxLength: 10 })
	.map((parts) => parts.join(' ').replace(/[\r\n]/g, ''));

/** Module text with every kind of line break the lexer accepts, mixed. */
const vbaSource = fc
	.array(fc.tuple(vbaLine, fc.constantFrom('\r\n', '\n', '\r')), { maxLength: 30 })
	.map((lines) => lines.map(([line, eol]) => line + eol).join(''));

describe('VBA lexer', () => {
	it('reproduces the source exactly from its tokens', () => {
		fc.assert(fc.property(vbaSource, (source) => {
			const tokens = tokenize(source);
			// Input that is trivia alone (whitespace, a bare ` _`) has no token to
			// carry it, and the lexer returns none (tokenize.ts, at end of input).
			if (tokens.length === 0) {
				return;
			}
			let text = '';
			for (const token of tokens) {
				for (const trivia of token.leadingTrivia ?? []) {
					text += trivia.text;
				}
				text += token.rawText;
				for (const trivia of token.trailingTrivia ?? []) {
					text += trivia.text;
				}
			}
			if (text !== source) {
				throw new Error(`tokens spell ${JSON.stringify(text)}`);
			}
		}), propertySettings());
	});
});

describe('VBA parser', () => {
	it('parses any module without throwing', () => {
		fc.assert(fc.property(vbaSource, (source) => {
			parseModule(source);
		}), propertySettings());
	});
});

describe('VBA formatter', () => {
	const options = fc.record({ tabSize: fc.integer({ min: 0, max: 8 }), insertSpaces: fc.boolean() });

	it('formats formatted text to itself', () => {
		fc.assert(fc.property(vbaSource, options, (source, chosen) => {
			const once = formatVbaModule(source, chosen).text;
			if (once === undefined) {
				return;
			}
			const twice = formatVbaModule(once, chosen);
			if (twice.text !== once) {
				throw new Error(`a second format gave ${JSON.stringify(twice.text)} (${twice.refusal ?? 'no refusal'}) for ${JSON.stringify(once)}`);
			}
		}), propertySettings());
	});
});

describe('VB6 form header', () => {
	const word = fc.stringMatching(/^[A-Za-z_][A-Za-z0-9_]{0,12}$/);
	const value = fc.oneof(
		fc.integer().map(String),
		fc.string({ maxLength: 12 }).map((text) => `"${text.replace(/["\r\n]/g, '')}"`),
		fc.constantFrom('-1  \'True', '"Form1.frx":0000', '&H8000000F&', '0'),
	);
	const property = fc.tuple(word, value).map(([key, text]) => `${key} = ${text}`);
	type Block = string[];
	const { block } = fc.letrec<{ block: Block; item: Block }>((tie) => ({
		item: fc.oneof(
			{ weight: 4, arbitrary: property.map((line) => [line]) },
			{ weight: 1, arbitrary: fc.tuple(word, fc.array(property, { maxLength: 3 })).map(([name, lines]) => [`BeginProperty ${name}`, ...lines, 'EndProperty']) },
			{ weight: 1, arbitrary: fc.tuple(word, word, tie('block')).map(([kind, name, body]) => [`Begin VB.${kind} ${name}`, ...body, 'End']) },
		),
		block: fc.array(tie('item'), { maxLength: 4 }).map((items) => items.flat()),
	}));
	const indent = (lines: string[]): string[] => {
		let depth = 0;
		return lines.map((line) => {
			if (/^(End|EndProperty)$/.test(line)) { depth = Math.max(0, depth - 1); }
			const out = '   '.repeat(depth) + line;
			if (/^Begin/.test(line)) { depth += 1; }
			return out;
		});
	};
	const frm = fc.tuple(
		word, block, fc.array(fc.constantFrom('Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"'), { maxLength: 2 }),
		fc.constantFrom('\r\n', '\n'),
		fc.array(fc.oneof(vbaLine, fc.constant('Attribute VB_Name = "Form1"')), { maxLength: 4 }),
		fc.option(fc.integer({ min: 0, max: 400 }), { nil: undefined }),
		fc.string({ minLength: 1, maxLength: 1 }),
	).map(([name, body, objects, eol, code, cut, junk]) => {
		const text = [
			'VERSION 5.00', ...objects, ...indent([`Begin VB.Form ${name}`, ...body, 'End']), ...code,
		].join(eol) + eol;
		// Some forms are damaged: a character changed somewhere in the header.
		// Line breaks are left alone: VB6 writes CRLF throughout, and the header
		// is printed with one line ending, the file's.
		const damageable = cut !== undefined && cut < text.length && !/[\r\n]/.test(text[cut] + junk);
		return damageable ? text.slice(0, cut) + junk + text.slice(cut + 1) : text;
	});

	it('reads a header back exactly as it was written, or refuses it', () => {
		fc.assert(fc.property(frm, (text) => {
			let header;
			try {
				header = parseFrmHeader(text);
			} catch (error) {
				if (error instanceof FrmHeaderError) {
					return;
				}
				throw error;
			}
			if (header === undefined) {
				return;
			}
			const printed = printFrmHeader(header);
			if (printed !== text.slice(0, header.endOffset)) {
				throw new Error(`printed ${JSON.stringify(printed)} for ${JSON.stringify(text.slice(0, header.endOffset))}`);
			}
		}), propertySettings());
	});
});

describe('VB6 project manifest', () => {
	const line = fc.oneof(
		fc.constantFrom(
			'Type=Exe', 'Form=Form1.frm', 'Module=Module1; Module1.bas', 'Class=Class1; Class1.cls',
			'Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\stdole2.tlb#OLE Automation',
			'Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX', 'Startup="Form1"',
			'[MS Transaction Server]', 'AutoRefresh=1', '', 'Name="Project1"',
		),
		fc.string({ maxLength: 30 }).map((text) => text.replace(/[\r\n]/g, '')),
	);
	const manifest = fc.tuple(fc.array(line, { maxLength: 15 }), fc.constantFrom('\r\n', '\n'), fc.boolean())
		.map(([lines, eol, trailing]) => lines.join(eol) + (trailing ? eol : ''));

	it('prints a manifest back byte for byte, or refuses it', () => {
		fc.assert(fc.property(manifest, (text) => {
			let parsed;
			try {
				parsed = parseVbpManifest(text);
			} catch (error) {
				if (error instanceof VbpManifestError) {
					return;
				}
				throw error;
			}
			const printed = printVbpManifest(parsed);
			if (printed !== text) {
				throw new Error(`printed ${JSON.stringify(printed)} for ${JSON.stringify(text)}`);
			}
		}), propertySettings());
	});
});
