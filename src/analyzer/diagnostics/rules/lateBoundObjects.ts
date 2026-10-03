// Rule family: late-bound objects whose state the code's literals make plain
// (issue #477). Measured in Excel 16.0 (build 20430, 2026-10-02); each
// compiles and raises when it runs. A ProgID no class has, a RegExp member
// or pattern VBScript refuses, Null into Test or Execute, and text into a
// Boolean flag are lateBoundMembers.ts's; this follows what the objects
// hold from one statement to the next.
//
//  - A VBScript.RegExp's Execute(s)(n) past the matches, or SubMatches(k)
//    past the groups, raises 5.
//  - A Scripting.FileSystemObject's OpenTextFile with an IOMode other than
//    1, 2 or 8 raises 5. A TextStream opened to read refuses a write, and
//    one opened to write or append refuses a read, with 54; after its Close
//    any member but Close raises 91. Through one path variable or literal,
//    the procedure's own file operations are followed: a file it created
//    and wrote nothing to raises 62 when read, CreateTextFile with
//    Overwrite False on it 58, and DeleteFile, GetFile or OpenTextFile to
//    read or append after it deleted the file 53. `If fso.FileExists(p)
//    Then fso.DeleteFile p` leaves no p.
//  - An ADODB.Recordset never opened raises 3704 at MoveNext, MoveFirst,
//    EOF, BOF, Close or RecordCount. State and Fields.Count run.
//  - An MSXML2.DOMDocument with no element, new or loaded from malformed
//    XML, raises 91 at DocumentElement's members, and SelectSingleNode of a
//    path no element of literal XML is on gives Nothing, whose members
//    raise 91. A path with a step and no node test raises -2147467259.
//
// A local is followed from the Set that creates it; any other mention of it
// whole, which may hand it to code that changes it, ends what is known.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { walkEnteringBlocks } from '../dataflow';
import { stringLiteralValue } from '../typeInference';
import { activeModuleMembers, matchParenFrom, setAssignmentTarget, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import { namesIn } from './shared';
import { foldStringExpression } from '../knownStringCalls';

/** The text constants an Execute subject is built with (issue #685). */
const SUBJECT_CONSTANTS: Readonly<Record<string, string>> = {
	vbcr: String.fromCharCode(13), vblf: String.fromCharCode(10), vbcrlf: String.fromCharCode(13, 10), vbnewline: String.fromCharCode(13, 10), vbtab: String.fromCharCode(9),
};

type LateObject =
	| { kind: 'regexp'; pattern: string | undefined; flags: { global?: boolean; ignorecase?: boolean; multiline?: boolean } }
	/** What the procedure has done to each path, by the variable or literal that names it. */
	| { kind: 'fso'; files: Record<string, FileFact> }
	/** `empty`: opened to read a file the procedure left empty. `file`: the path it writes. */
	| { kind: 'textstream'; mode: 'read' | 'write'; closed: boolean; empty?: boolean; file?: string }
	| { kind: 'recordset'; open?: boolean; closed?: boolean }
	/** The document element, or null for a document with none: new, or loaded from malformed XML. */
	| { kind: 'domdoc'; root: XmlElement | null };

/** A path the procedure created empty, wrote to, or deleted. */
type FileFact = 'empty' | 'full' | 'absent';

interface XmlElement {
	name: string;
	children: XmlElement[];
	/** Each attribute's value as written; undefined where it holds an entity reference. */
	attributes?: Readonly<Record<string, string | undefined>>;
}

const REGEXP_MEMBERS: ReadonlySet<string> = new Set(['pattern', 'global', 'ignorecase', 'multiline', 'test', 'execute', 'replace']);
const REGEXP_FLAGS: ReadonlySet<string> = new Set(['global', 'ignorecase', 'multiline']);
const READS: ReadonlySet<string> = new Set(['readline', 'readall', 'read', 'skip', 'skipline']);
const WRITES: ReadonlySet<string> = new Set(['write', 'writeline', 'writeblanklines']);
const CLOSED_RECORDSET: ReadonlySet<string> = new Set(['movenext', 'moveprevious', 'movefirst', 'movelast', 'eof', 'bof', 'close', 'recordcount']);

export function checkLateBoundObjects(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const states = new Map<string, LateObject>();
		const forget = (names: Iterable<string>): void => {
			for (const lower of names) {
				states.delete(lower);
			}
		};
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				states.clear();
			}
			// `p = ...` names another path with the same variable.
			const head = tokenText(toks[0]) === 'let' ? 1 : 0;
			const assigned = toks[head + 1]?.rawText === '=' ? tokenName(toks[head])?.toLowerCase() : undefined;
			if (assigned) {
				for (const state of states.values()) {
					if (state.kind === 'fso') {
						delete state.files[assigned];
					}
				}
			}
			// A one-line If: its condition and each arm are checked on a copy,
			// and what an arm may change is no longer known after it.
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				const then = toks.findIndex((tok) => tokenText(tok) === 'then');
				checkMembers(node.span, toks.slice(1, then > 0 ? then : toks.length), states, push);
				for (const branch of node.singleLineIfBranches) {
					if (setAssignmentTarget(source, branch)) {
						forget(namesIn(source, branch));
						continue;
					}
					const armToks = statementTokensAfterLeadingLabel(source, branch).filter((tok) => tok.kind !== 'comment' && tokenText(tok) !== 'else');
					// `If fso.FileExists(p) Then fso.DeleteFile p` deletes only what
					// is there, and leaves no p either way.
					const guarded = [...states].find(([lower, state]) => state.kind === 'fso' && deletesWhatItTests(toks.slice(1, then), armToks, lower) !== undefined);
					if (guarded && guarded[1].kind === 'fso') {
						guarded[1].files[deletesWhatItTests(toks.slice(1, then), armToks, guarded[0])!] = 'absent';
						continue;
					}
					const arm = new Map([...states].map(([lower, state]) => [lower, copy(state)]));
					checkMembers(branch, armToks, arm, push);
					for (const [lower, state] of [...states]) {
						const after = arm.get(lower);
						if (state.kind === 'fso' && after?.kind === 'fso') {
							// A path the arm may change is unknown after it.
							for (const file of new Set([...Object.keys(state.files), ...Object.keys(after.files)])) {
								if (state.files[file] !== after.files[file]) {
									delete state.files[file];
								}
							}
						} else if (JSON.stringify(after) !== JSON.stringify(state)) {
							states.delete(lower);
						}
					}
				}
				return;
			}
			const set = setAssignmentTarget(source, node.span);
			if (set) {
				const lower = set.name.toLowerCase();
				const eq = toks.findIndex((tok) => tok.rawText === '=');
				const value = toks.slice(eq + 1);
				const created = createdObject(value, states);
				checkMembers(node.span, toks.slice(eq + 1), states, push);
				// `Set ts = fso.CreateTextFile(p)` reads fso, which stays followed
				// with what the call did to its files (issue #685).
				const named = namesIn(source, node.span);
				if (created?.kind === 'textstream') {
					named.delete(tokenName(value[0])?.toLowerCase() ?? '');
				}
				forget(named);
				if (created) {
					states.set(lower, created);
				}
				return;
			}
			checkMembers(node.span, toks, states, push);
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map([...states].map(([lower, state]) => [lower, copy(state)])),
			restore: (saved) => {
				states.clear();
				for (const [lower, state] of saved) {
					states.set(lower, copy(state));
				}
			},
			forget,
			touches: (stmt) => namesIn(source, stmt.span),
		});
	}
}

function copy(state: LateObject): LateObject {
	// A document's element tree is never changed in place, so it is shared.
	return state.kind === 'regexp' ? { ...state, flags: { ...state.flags } }
		: state.kind === 'fso' ? { kind: 'fso', files: { ...state.files } }
		: { ...state };
}

/** What `CreateObject("...")` or `fso.OpenTextFile(...)` gives, to be followed under the Set's target. */
function createdObject(value: readonly VbaToken[], states: ReadonlyMap<string, LateObject>): LateObject | undefined {
	const progId = createObjectProgId(value);
	if (progId !== undefined) {
		const id = progId.toLowerCase();
		return id === 'vbscript.regexp' ? { kind: 'regexp', pattern: '', flags: { global: false, ignorecase: false, multiline: false } }
			: id === 'scripting.filesystemobject' ? { kind: 'fso', files: {} }
			: id === 'adodb.recordset' ? { kind: 'recordset' }
			: /^msxml2\.domdocument(\.\d+\.\d+)?$/.test(id) || id === 'microsoft.xmldom' ? { kind: 'domdoc', root: null }
			: undefined;
	}
	// `fso.CreateTextFile(p)` writes; `fso.OpenTextFile(p, mode)` reads with
	// mode 1 or none, and writes with 2 or 8.
	const owner = tokenName(value[0])?.toLowerCase();
	const method = tokenText(value[2]);
	const fso = owner ? states.get(owner) : undefined;
	if (fso?.kind !== 'fso' || value[1]?.rawText !== '.' || value[3]?.rawText !== '(' || matchParenFrom(value, 3) !== value.length - 1) {
		return undefined;
	}
	const file = pathKey(splitTopLevelTokenGroups(value, 4, ',', value.length - 1)[0]);
	// A stream to write leaves the file empty until something is written:
	// `Set ts = fso.CreateTextFile(p): ts.Close`, then reading p raises 62
	// (issue #685, measured in Excel 16.0). Append keeps what it held.
	if (method === 'createtextfile') {
		if (file !== undefined) {
			fso.files[file] = 'empty';
		}
		return { kind: 'textstream', mode: 'write', closed: false, file };
	}
	if (method === 'opentextfile') {
		const mode = ioMode(value, 3);
		if (mode === 2 && file !== undefined) {
			fso.files[file] = 'empty';
		}
		return mode === 'none' || mode === 1 ? { kind: 'textstream', mode: 'read', closed: false, empty: file !== undefined && fso.files[file] === 'empty' }
			: mode === 2 || mode === 8 ? { kind: 'textstream', mode: 'write', closed: false, file }
			: undefined;
	}
	return undefined;
}

/** The path `fso.FileExists(p)` tests when the arm is `fso.DeleteFile p` on the same path. */
function deletesWhatItTests(condition: readonly VbaToken[], arm: readonly VbaToken[], fso: string): string | undefined {
	const tested = condition.length >= 6 && tokenName(condition[0])?.toLowerCase() === fso && condition[1].rawText === '.' && tokenText(condition[2]) === 'fileexists'
		&& condition[3].rawText === '(' && matchParenFrom(condition, 3) === condition.length - 1
		? pathKey(condition.slice(4, condition.length - 1)) : undefined;
	const deleted = arm.length === 4 && tokenName(arm[0])?.toLowerCase() === fso && arm[1].rawText === '.' && tokenText(arm[2]) === 'deletefile'
		? pathKey([arm[3]]) : undefined;
	return tested !== undefined && tested === deleted ? tested : undefined;
}

/** The key a path argument is known by: a variable's name, or a literal's text. */
function pathKey(arg: readonly VbaToken[] | undefined): string | undefined {
	const toks = arg?.filter((tok) => tok.kind !== 'comment');
	if (toks?.length !== 1) {
		return undefined;
	}
	return toks[0].kind === 'stringLiteral' ? toks[0].rawText : tokenName(toks[0])?.toLowerCase();
}

/** The IOMode an `OpenTextFile(` at `open` passes: a literal, 'none' when omitted, undefined otherwise. */
function ioMode(toks: readonly VbaToken[], open: number): number | 'none' | undefined {
	const args = splitTopLevelTokenGroups(toks, open + 1, ',', matchParenFrom(toks, open));
	const arg = args[1]?.filter((tok) => tok.kind !== 'comment');
	if (!arg || arg.length === 0) {
		return 'none';
	}
	return arg.length === 1 && arg[0].kind === 'integerLiteral' && /^\d+$/.test(arg[0].rawText) ? Number(arg[0].rawText) : undefined;
}

/** The ProgID of a whole `CreateObject("...")` value. */
function createObjectProgId(value: readonly VbaToken[]): string | undefined {
	const at = tokenText(value[0]) === 'vba' && value[1]?.rawText === '.' ? 2 : 0;
	if (tokenText(value[at]) !== 'createobject' || value[at + 1]?.rawText !== '(' || value[at + 2]?.kind !== 'stringLiteral'
		|| (value[at + 3]?.rawText !== ')' && value[at + 3]?.rawText !== ',') || matchParenFrom(value, at + 1) !== value.length - 1) {
		return undefined;
	}
	return stringLiteralValue(value[at + 2].rawText);
}

/** Each `x.Member` on a followed object; any other mention of x ends what is known of it. */
function checkMembers(base: Span, toks: readonly VbaToken[], states: Map<string, LateObject>, push: PushFn): void {
	const at = (tok: VbaToken): Span => ({ start: base.start + tok.start, end: base.start + tok.end });
	const ended = new Set<string>();
	for (let i = 0; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const state = lower ? states.get(lower) : undefined;
		if (!state || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		if (toks[i + 1]?.rawText !== '.' || !tokenName(toks[i + 2])) {
			ended.add(lower!);
			continue;
		}
		const memberTok = toks[i + 2];
		const memberName = tokenText(memberTok);
		const assigned = i === 0 && toks[3]?.rawText === '=' ? toks.slice(4) : undefined;
		if (state.kind === 'textstream') {
			if (state.closed && memberName !== 'close') {
				push('objectVariableNotSet', `'${toks[i].rawText}' was closed, so its ${memberTok.rawText} has no stream to reach. This will raise Run-time error '91': Object variable or With block variable not set.`, at(memberTok));
				ended.add(lower!);
			} else if ((state.mode === 'write' && READS.has(memberName)) || (state.mode === 'read' && WRITES.has(memberName))) {
				push('fileModeMismatch', `'${toks[i].rawText}' was opened to ${state.mode === 'read' ? 'read' : 'write'}, so ${memberTok.rawText} cannot use it. This will raise Run-time error '54': Bad file mode.`, at(memberTok));
			} else if (state.empty && READS.has(memberName)) {
				push('fileModeMismatch', `'${toks[i].rawText}' reads a file this procedure left empty, so ${memberTok.rawText} reads past its end. This will raise Run-time error '62': Input past end of file.`, at(memberTok));
			} else if (WRITES.has(memberName) && state.file !== undefined) {
				for (const other of states.values()) {
					if (other.kind === 'fso') {
						other.files[state.file] = 'full';
					}
				}
			} else if (memberName === 'close') {
				state.closed = true;
			}
			continue;
		}
		if (state.kind === 'recordset') {
			// Open, then Close, leaves it closed again: a second Close raises
			// 3704 too (issue #685, measured in Excel 16.0).
			if (memberName === 'open') {
				state.open = true;
			} else if (state.open && memberName === 'close') {
				state.open = false;
				state.closed = true;
			} else if (!state.open && CLOSED_RECORDSET.has(memberName)) {
				push('lateBoundObjectState', `'${toks[i].rawText}' ${state.closed ? 'was closed above' : 'was never opened'}, so ${memberTok.rawText} has no records to work on. This will raise Run-time error '3704': Operation is not allowed when the object is closed.`, at(memberTok));
			}
			continue;
		}
		if (state.kind === 'fso') {
			fileOperation(toks, i, state, at, push);
			continue;
		}
		if (state.kind === 'domdoc') {
			checkDocument(toks, i, state, at, push, () => ended.add(lower!));
			continue;
		}
		// A RegExp. Its members, patterns and flags are runtime-member-not-found's
		// and runtime-argument-value's; this follows what Execute finds.
		if (!REGEXP_MEMBERS.has(memberName)) {
			continue;
		}
		if (assigned) {
			const value = assigned.length === 1 ? assigned[0] : undefined;
			if (memberName === 'pattern') {
				state.pattern = value?.kind === 'stringLiteral' ? stringLiteralValue(value.rawText) : undefined;
			} else if (REGEXP_FLAGS.has(memberName)) {
				const flag = memberName as 'global' | 'ignorecase' | 'multiline';
				const text = value?.kind === 'stringLiteral' ? stringLiteralValue(value.rawText).trim().toLowerCase() : undefined;
				const word = value ? tokenText(value) : '';
				state.flags[flag] = word === 'true' || text === 'true' ? true : word === 'false' || text === 'false' ? false : undefined;
			}
			break;
		}
		if ((memberName === 'test' || memberName === 'execute' || memberName === 'replace') && toks[i + 3]?.rawText === '(') {
			if (state.pattern === undefined || vbscriptPatternError(state.pattern) !== undefined) {
				continue;
			}
			const close = matchParenFrom(toks, i + 3);
			const args = splitTopLevelTokenGroups(toks, i + 4, ',', close);
			// The subject as a literal, or literals and vbCr, vbLf, vbCrLf and
			// vbTab joined by &: `"x" & vbLf & "x"` (issue #685).
			const subject = memberName === 'execute' && args.length === 1
				? foldStringExpression(args[0], { nameValue: (tok) => SUBJECT_CONSTANTS[tokenText(tok)], integerValue: () => undefined })
				: undefined;
			if (memberName === 'execute' && state.pattern !== undefined && subject !== undefined && !(state.flags.multiline && subject.includes('\r'))) {
				const hit = matchIndexFault(state.pattern, state.flags, subject, toks, close);
				if (hit) {
					push('collectionIndexOutOfRange', hit.message, at(toks[hit.at]));
				}
			}
		}
	}
	for (const lower of ended) {
		states.delete(lower);
	}
}

/**
 * `Execute(s)(n)` past the matches, or `(n).SubMatches(k)` past the
 * groups, from the token after Execute's `)`. Matched as JavaScript does,
 * which VBScript's patterns share for what the checker lets through.
 */
function matchIndexFault(
	pattern: string,
	flags: { global?: boolean; ignorecase?: boolean; multiline?: boolean },
	subject: string,
	toks: readonly VbaToken[],
	close: number,
): { at: number; message: string } | undefined {
	const index = literalIndexAt(toks, close + 1);
	if (index === undefined || flags.global === undefined || flags.ignorecase === undefined || flags.multiline === undefined) {
		return undefined;
	}
	let expression: RegExp;
	try {
		expression = new RegExp(pattern, `g${flags.ignorecase ? 'i' : ''}${flags.multiline ? 'm' : ''}`);
	} catch {
		return undefined;
	}
	const matches: RegExpExecArray[] = [];
	for (let found = expression.exec(subject); found; found = expression.exec(subject)) {
		if (found[0] === '') {
			return undefined;
		}
		matches.push(found);
		if (!flags.global) {
			break;
		}
	}
	if (index.value >= matches.length) {
		return { at: close + 2, message: `The pattern finds ${matches.length === 0 ? 'no match' : `${matches.length} match${matches.length === 1 ? '' : 'es'}`} in ${JSON.stringify(subject)}, so match ${index.value} is past them. This will raise Run-time error '5': Invalid procedure call or argument.` };
	}
	const after = index.end + 1;
	if (toks[after]?.rawText === '.' && tokenText(toks[after + 1]) === 'submatches') {
		const group = literalIndexAt(toks, after + 2);
		const groups = matches[index.value].length - 1;
		if (group && group.value >= groups) {
			return { at: after + 3, message: `The pattern has ${groups} group${groups === 1 ? '' : 's'}, so SubMatches(${group.value}) is past them. This will raise Run-time error '5': Invalid procedure call or argument.` };
		}
	}
	return undefined;
}

/** FileSystemObject members that take a path first and change no file. */
const FILE_READS: ReadonlySet<string> = new Set(['fileexists', 'folderexists', 'getfilename', 'getbasename', 'getextensionname', 'getparentfoldername', 'getabsolutepathname', 'buildpath', 'gettempname', 'getspecialfolder', 'drives']);

/**
 * What a FileSystemObject call does to the path it names first (issue #477,
 * measured in Excel 16.0): CreateTextFile with Overwrite False on a file the
 * procedure made raises 58; DeleteFile, GetFile and OpenTextFile to read or
 * append without Create on a file it deleted raise 53. CreateTextFile, and
 * OpenTextFile to write, leave the file empty; DeleteFile leaves none. Any
 * other call may change any file.
 */
function fileOperation(
	toks: readonly VbaToken[],
	i: number,
	state: { kind: 'fso'; files: Record<string, FileFact> },
	at: (tok: VbaToken) => Span,
	push: PushFn,
): void {
	const memberTok = toks[i + 2];
	const member = tokenText(memberTok);
	if (FILE_READS.has(member)) {
		return;
	}
	if (member === 'opentextfile' && toks[i + 3]?.rawText === '(') {
		const mode = ioMode(toks, i + 3);
		if (typeof mode === 'number' && mode !== 1 && mode !== 2 && mode !== 8) {
			const arg = splitTopLevelTokenGroups(toks, i + 4, ',', matchParenFrom(toks, i + 3))[1];
			push('runtimeArgumentValue', `OpenTextFile's IOMode is 1 to read, 2 to write or 8 to append, and ${mode} is none of them. This will raise Run-time error '5': Invalid procedure call or argument.`, at(arg[0]));
		}
	}
	const parens = toks[i + 3]?.rawText === '(';
	const end = parens ? matchParenFrom(toks, i + 3) : toks.length;
	const first = parens ? i + 4 : i + 3;
	const args = end > first ? splitTopLevelTokenGroups(toks, first, ',', end) : [];
	const file = pathKey(args[0]);
	if (file === undefined || !['createtextfile', 'opentextfile', 'deletefile', 'getfile'].includes(member)) {
		state.files = {};
		return;
	}
	const fact = state.files[file];
	const shown = args[0][0].rawText;
	const literal = (k: number): string | undefined => {
		const arg = args[k]?.filter((tok) => tok.kind !== 'comment');
		return arg?.length === 1 ? tokenText(arg[0]) || arg[0].rawText : undefined;
	};
	if (member === 'createtextfile') {
		if (literal(1) === 'false' && (fact === 'empty' || fact === 'full')) {
			push('runtimeArgumentValue', `${memberTok.rawText} with Overwrite False finds ${shown}, which this procedure made. This will raise Run-time error '58': File already exists.`, at(memberTok));
		}
		state.files[file] = 'empty';
		return;
	}
	if (member === 'opentextfile') {
		const mode = ioMode(toks, i + 3);
		const create = literal(2) === 'true';
		if (fact === 'absent' && (mode === 'none' || mode === 1 || (mode === 8 && !create))) {
			push('runtimeArgumentValue', `${memberTok.rawText} finds no ${shown}, which this procedure deleted. This will raise Run-time error '53': File not found.`, at(memberTok));
		}
		if (mode === 2 && (create || fact !== 'absent')) {
			state.files[file] = 'empty';
		} else if (mode === 8 && create && fact === 'absent') {
			state.files[file] = 'empty';
		}
		return;
	}
	if (fact === 'absent') {
		push('runtimeArgumentValue', `${memberTok.rawText} finds no ${shown}, which this procedure deleted. This will raise Run-time error '53': File not found.`, at(memberTok));
	}
	if (member === 'deletefile') {
		state.files[file] = 'absent';
	}
}

/** Members of a DOMDocument that read it and change nothing. */
const DOCUMENT_READS: ReadonlySet<string> = new Set(['documentelement', 'selectsinglenode', 'selectnodes', 'xml', 'text', 'parseerror', 'getelementsbytagname']);

/**
 * `doc.LoadXML "<a>"` and what reads the document after it: a document with
 * no element, new or loaded from malformed XML, has a DocumentElement of
 * Nothing; SelectSingleNode of a path no element is on gives Nothing; and
 * a path with a step and no node test raises -2147467259 (issue #477,
 * measured in Excel 16.0). Anything else that may change the document ends
 * what is known of it.
 */
function checkDocument(
	toks: readonly VbaToken[],
	i: number,
	state: { kind: 'domdoc'; root: XmlElement | null },
	at: (tok: VbaToken) => Span,
	push: PushFn,
	end: () => void,
): void {
	const memberTok = toks[i + 2];
	const memberName = tokenText(memberTok);
	if (memberName === 'loadxml') {
		const arg = toks[i + 3]?.rawText === '(' && toks[i + 5]?.rawText === ')' ? toks[i + 4]
			: i === 0 && toks.length === 4 ? toks[3]
			: undefined;
		const parsed = arg?.kind === 'stringLiteral' ? parseXml(stringLiteralValue(arg.rawText)) : undefined;
		if (parsed === undefined) {
			end();
		} else {
			state.root = parsed;
		}
		return;
	}
	if (!DOCUMENT_READS.has(memberName)) {
		end();
		return;
	}
	const reached = (after: number): VbaToken | undefined => (toks[after]?.rawText === '.' && tokenName(toks[after + 1]) ? toks[after + 1] : undefined);
	if (memberName === 'documentelement' && state.root === null) {
		const next = reached(i + 3);
		if (next) {
			push('objectVariableNotSet', `'${toks[i].rawText}' holds no element, as it is new or was loaded from malformed XML, so its DocumentElement is Nothing and '.${next.rawText}' has no object to reach. This will raise Run-time error '91': Object variable or With block variable not set.`, at(next));
		}
		return;
	}
	if ((memberName === 'selectsinglenode' || memberName === 'selectnodes') && toks[i + 3]?.rawText === '(' && toks[i + 4]?.kind === 'stringLiteral' && toks[i + 5]?.rawText === ')') {
		const path = stringLiteralValue(toks[i + 4].rawText);
		if (/(^|\/)\[/.test(path)) {
			push('runtimeArgumentValue', `The XPath '${path}' has a step with no node test before its '['. This will raise Run-time error '-2147467259': NodeTest expected here.`, at(toks[i + 4]));
			return;
		}
		const next = memberName === 'selectsinglenode' ? reached(i + 6) : undefined;
		if (next && (state.root === null || pathFinds(state.root, path) === false)) {
			push('objectVariableNotSet', `No element of '${toks[i].rawText}' is on the path '${path}', so SelectSingleNode gives Nothing and '.${next.rawText}' has no object to reach. This will raise Run-time error '91': Object variable or With block variable not set.`, at(next));
		}
	}
}

/**
 * Whether an element is on `//name`, `/a/b` or `a/b`, each step with an
 * optional `[n]` or `[@attr='value']` predicate: `//b[3]`, the third b of
 * some parent, and `//b[@id='2']` (issue #685, measured in Excel 16.0 with
 * MSXML 6). Undefined for any other path. Names match as written.
 */
function pathFinds(root: XmlElement, path: string): boolean | undefined {
	const name = '[A-Za-z_][\\w.-]*';
	const step = `${name}(?:\\[(?:\\d+|@${name}\\s*=\\s*(?:'[^']*'|"[^"]*"))\\])?`;
	const descendant = new RegExp(`^//${step}$`).test(path);
	if (!descendant && !new RegExp(`^/?${step}(/${step})*$`).test(path)) {
		return undefined;
	}
	// What a step keeps of one parent's children: undefined where an attribute it reads is not known.
	const take = (siblings: readonly XmlElement[], text: string): XmlElement[] | undefined => {
		const parts = /^([^[]+)(?:\[(?:(\d+)|@([^=\s]+)\s*=\s*(?:'([^']*)'|"([^"]*)"))\])?$/.exec(text)!;
		const named = siblings.filter((element) => element.name === parts[1]);
		if (parts[2] !== undefined) {
			const at = named[Number(parts[2]) - 1];
			return at ? [at] : [];
		}
		if (parts[3] === undefined) {
			return named;
		}
		const wanted = parts[4] ?? parts[5];
		if (named.some((element) => element.attributes === undefined || (parts[3] in element.attributes && element.attributes[parts[3]] === undefined))) {
			return undefined;
		}
		return named.filter((element) => element.attributes![parts[3]] === wanted);
	};
	if (descendant) {
		// Every parent's children, the document's own (the root) included.
		const groups: XmlElement[][] = [[root]];
		const gather = (element: XmlElement): void => {
			groups.push(element.children);
			element.children.forEach(gather);
		};
		gather(root);
		let unknown = false;
		for (const group of groups) {
			const kept = take(group, path.slice(2));
			if (kept === undefined) {
				unknown = true;
			} else if (kept.length > 0) {
				return true;
			}
		}
		return unknown ? undefined : false;
	}
	const steps = path.replace(/^\//, '').match(new RegExp(step, 'g'))!;
	let level: readonly XmlElement[][] = [[root]];
	for (const text of steps) {
		const found: XmlElement[] = [];
		for (const group of level) {
			const kept = take(group, text);
			if (kept === undefined) {
				return undefined;
			}
			found.push(...kept);
		}
		if (found.length === 0) {
			return false;
		}
		level = found.map((element) => element.children);
	}
	return true;
}

/**
 * The document element of `text` as XML: null where it is malformed, and
 * undefined where this does not judge it (a DOCTYPE). Read strictly, as
 * MSXML's LoadXML does: quoted attributes, one root, entities only as
 * `&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;` or a character reference.
 */
export function parseXml(text: string): XmlElement | null | undefined {
	let i = 0;
	const name = /[A-Za-z_:][\w.:-]*/y;
	const space = (): void => {
		while (i < text.length && /\s/.test(text[i])) {
			i++;
		}
	};
	const readName = (): string | undefined => {
		name.lastIndex = i;
		const found = name.exec(text);
		if (!found) {
			return undefined;
		}
		i += found[0].length;
		return found[0];
	};
	const skipMisc = (): boolean | undefined => {
		for (;;) {
			space();
			if (text.startsWith('<!--', i)) {
				const close = text.indexOf('-->', i + 4);
				if (close < 0) {
					return false;
				}
				i = close + 3;
			} else if (text.startsWith('<?', i)) {
				const close = text.indexOf('?>', i + 2);
				if (close < 0) {
					return false;
				}
				i = close + 2;
			} else if (text.startsWith('<!DOCTYPE', i)) {
				return undefined;
			} else {
				return true;
			}
		}
	};
	const element = (): XmlElement | null => {
		if (text[i] !== '<') {
			return null;
		}
		i++;
		const tag = readName();
		if (!tag) {
			return null;
		}
		const seen = new Set<string>();
		const attributes: Record<string, string | undefined> = {};
		for (;;) {
			const before = i;
			space();
			if (text.startsWith('/>', i)) {
				i += 2;
				return { name: tag, children: [], attributes };
			}
			if (text[i] === '>') {
				i++;
				break;
			}
			if (i === before) {
				return null;
			}
			const attr = readName();
			space();
			if (!attr || seen.has(attr) || text[i] !== '=') {
				return null;
			}
			seen.add(attr);
			i++;
			space();
			const quote = text[i];
			if (quote !== '"' && quote !== "'") {
				return null;
			}
			const close = text.indexOf(quote, i + 1);
			if (close < 0 || text.slice(i + 1, close).includes('<') || !entitiesOk(text.slice(i + 1, close))) {
				return null;
			}
			const raw = text.slice(i + 1, close);
			attributes[attr] = raw.includes('&') ? undefined : raw;
			i = close + 1;
		}
		const children: XmlElement[] = [];
		for (;;) {
			const lt = text.indexOf('<', i);
			if (lt < 0 || !entitiesOk(text.slice(i, lt))) {
				return null;
			}
			i = lt;
			if (text.startsWith('</', i)) {
				i += 2;
				const closing = readName();
				space();
				if (closing !== tag || text[i] !== '>') {
					return null;
				}
				i++;
				return { name: tag, children, attributes };
			}
			if (text.startsWith('<!--', i)) {
				const close = text.indexOf('-->', i + 4);
				if (close < 0) {
					return null;
				}
				i = close + 3;
			} else if (text.startsWith('<![CDATA[', i)) {
				const close = text.indexOf(']]>', i + 9);
				if (close < 0) {
					return null;
				}
				i = close + 3;
			} else if (text.startsWith('<?', i)) {
				const close = text.indexOf('?>', i + 2);
				if (close < 0) {
					return null;
				}
				i = close + 2;
			} else {
				const child = element();
				if (!child) {
					return null;
				}
				children.push(child);
			}
		}
	};
	const before = skipMisc();
	if (before !== true) {
		return before === undefined ? undefined : null;
	}
	const root = element();
	if (!root) {
		return null;
	}
	const after = skipMisc();
	return after === undefined ? undefined : after && i === text.length ? root : null;
}

/** Whether every `&` in XML text starts an entity XML defines. */
function entitiesOk(text: string): boolean {
	return !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)/.test(text);
}

/** A `(n)` with n a whole literal at `open`: n and the index of its `)`. */
function literalIndexAt(toks: readonly VbaToken[], open: number): { value: number; end: number } | undefined {
	return toks[open]?.rawText === '(' && toks[open + 1]?.kind === 'integerLiteral' && /^\d+$/.test(toks[open + 1].rawText) && toks[open + 2]?.rawText === ')'
		? { value: Number(toks[open + 1].rawText), end: open + 2 }
		: undefined;
}

/**
 * The error VBScript's regular expressions raise for a pattern they cannot
 * read, or undefined: 5020 for an open group, 5019 for an open class, 5018
 * for a quantifier with nothing to repeat, 5017 otherwise (issue #477,
 * measured in Excel 16.0).
 */
export function vbscriptPatternError(pattern: string): number | undefined {
	let depth = 0;
	// Whether what came last can take a quantifier, and whether it was one.
	let atom = false;
	let quantified = false;
	// A `?` right after a quantifier makes it lazy, once: `a??` runs, `a+??` does not.
	let lazyable = false;
	for (let i = 0; i < pattern.length; i++) {
		const c = pattern[i];
		if (c === '?' && lazyable) {
			lazyable = false;
			continue;
		}
		lazyable = false;
		if (c === '\\') {
			if (i + 1 >= pattern.length) {
				return 5017;
			}
			i++;
			// `\b` and `\B` are anchors, which nothing repeats: `\b*` raises 5018.
			atom = pattern[i] !== 'b' && pattern[i] !== 'B';
			quantified = false;
			continue;
		}
		if (c === '[') {
			let j = i + 1;
			if (pattern[j] === '^') {
				j++;
			}
			// A `]` first in the class is one of its characters: `[]a]` runs,
			// and `[^]` is left open (5019).
			if (pattern[j] === ']') {
				j++;
			}
			for (; j < pattern.length && pattern[j] !== ']'; j++) {
				if (pattern[j] === '\\') {
					j++;
				}
			}
			if (j >= pattern.length) {
				return 5019;
			}
			i = j;
			atom = true;
			quantified = false;
			continue;
		}
		if (c === '(') {
			if (pattern[i + 1] === '?') {
				if (!':=!'.includes(pattern[i + 2] ?? 'x')) {
					return 5017;
				}
				i += 2;
			}
			depth++;
			atom = false;
			quantified = false;
			continue;
		}
		if (c === ')') {
			if (depth === 0) {
				return 5017;
			}
			depth--;
			atom = true;
			quantified = false;
			continue;
		}
		if (c === '|') {
			atom = false;
			quantified = false;
			continue;
		}
		const braces = c === '{' ? /^\{(\d+)(,(\d*))?\}/.exec(pattern.slice(i)) : null;
		if (c === '*' || c === '+' || c === '?' || braces) {
			if (!atom || quantified) {
				return 5018;
			}
			if (braces) {
				if (braces[3] !== undefined && braces[3] !== '' && Number(braces[3]) < Number(braces[1])) {
					return 5017;
				}
				i += braces[0].length - 1;
			}
			quantified = true;
			lazyable = true;
			continue;
		}
		atom = c !== '^' && c !== '$';
		quantified = false;
	}
	return depth > 0 ? 5020 : undefined;
}
