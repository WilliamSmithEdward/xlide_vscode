// Rule: AddressOf where the VBE refuses it (issue #299, each measured in
// 64-bit Excel 16.0 as a compile error).
//
// AddressOf compiles only as a whole argument of a project procedure or of
// a method such as Collection.Add, and only on a Sub, Function or Property
// of a standard module. Elsewhere:
//
//  - outside an argument, `p = AddressOf Cb`, in parentheses of its own,
//    `Take((AddressOf Cb))`, in Debug.Print, or as the argument of Len,
//    CLng, CLngPtr, CStr or Abs: "Syntax error"; of ObjPtr: "Type
//    mismatch". VarPtr, StrPtr, Hex, IsEmpty and TypeName take it. With an
//    operator after it, `Take(AddressOf Cb + 1)`: "Argument not optional";
//  - on a name nothing declares, or a class's member, `AddressOf K.Run`:
//    "Variable not defined" (under Option Explicit); on a variable:
//    "Expected Sub, Function, or Property"; on a VBA function such as Len:
//    "Syntax error"; on a Declare: "Invalid use of AddressOf operator";
//  - into a ByVal Long, Integer or Byte parameter in 64-bit VBA: "Type
//    mismatch", since AddressOf gives a LongPtr (#298).

import {
	type ConditionalActivityTracker,
	type ConditionalCompilationEnvironment,
	compilerConstantsWithDefaults,
} from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import { resolveRuntimeFunction } from '../../runtime/vbaRuntime';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProcedureSignature, VbaProjectClassMembers } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { VALUE_WORD_ERRORS } from './malformedLines';
import { callableTypeSignaturesFor, normalizeType } from '../typeInference';
import {
	activeModuleMembers,
	forEachStatement,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

const NARROW: ReadonlySet<string> = new Set(['long', 'integer', 'byte']);

/** The VBA functions measured to refuse an AddressOf argument, and the error each gives. */
const ADDRESSOF_REFUSED_BY: ReadonlyMap<string, string> = new Map([
	['len', 'Syntax error'],
	['clng', 'Syntax error'],
	['clngptr', 'Syntax error'],
	['cstr', 'Syntax error'],
	['abs', 'Syntax error'],
	['objptr', 'Type mismatch'],
]);

export function checkAddressOfUse(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectClassMembers: readonly VbaProjectClassMembers[] | undefined,
	conditionalCompilation: ConditionalCompilationEnvironment | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!/\bAddressOf\b/i.test(source)) {
		return;
	}
	const win64 = compilerConstantsWithDefaults(conditionalCompilation).get('win64');
	const is64 = typeof win64 === 'number' ? win64 !== 0 : win64 === true;
	const explicit = /^[ \t]*Option[ \t]+Explicit\b/im.test(source);
	const signatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const moduleMembers = new Map((symbols.root.children ?? []).map((child) => [child.name.toLowerCase(), child.kind]));
	const standardModules = new Map((projectClassMembers ?? []).filter((type) => type.kind === 'standardModule').map((type) => [type.name.toLowerCase(), new Set(type.members.map((member) => member.name.toLowerCase()))]));
	const projectProcedureNames = new Set([...standardModules.values()].flatMap((names) => [...names]));
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = new Set([
			...member.params.map((param) => param.name.toLowerCase()),
			...(procedureSymbolFor(symbols, member)?.children ?? []).map((child) => child.name.toLowerCase()),
		]);
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				toks.forEach((tok, i) => {
					if (tokenText(tok) !== 'addressof') {
						return;
					}
					const problem = addressOfProblem(toks, i, {
						member: member.name.toLowerCase(),
						locals,
						moduleMembers,
						standardModules,
						projectProcedureNames,
						signatures,
						is64,
						explicit,
					});
					if (problem) {
						const end = toks[i + 1] && tokenName(toks[i + 1]) ? (toks[i + 2]?.rawText === '.' && toks[i + 3] ? i + 3 : i + 1) : i;
						const at: Span = { start: span.start + tok.start, end: span.start + toks[end].end };
						push('addressOfMisuse', `${problem.why} This is a VBE compile error: ${problem.error}.`, at);
					}
				});
			}
		}, activity);
	}
}

interface Scope {
	member: string;
	locals: ReadonlySet<string>;
	moduleMembers: ReadonlyMap<string, string>;
	standardModules: ReadonlyMap<string, ReadonlySet<string>>;
	projectProcedureNames: ReadonlySet<string>;
	signatures: ReturnType<typeof callableTypeSignaturesFor>;
	is64: boolean;
	explicit: boolean;
}

function addressOfProblem(toks: readonly VbaToken[], at: number, scope: Scope): { why: string; error: string } | undefined {
	const first = toks[at + 1];
	const name = tokenName(first);
	if (!name) {
		return undefined;
	}
	const qualified = toks[at + 2]?.rawText === '.' && tokenName(toks[at + 3]) !== undefined;
	const last = qualified ? at + 3 : at + 1;
	const text = toks.slice(at + 1, last + 1).map((tok) => tok.rawText).join('');
	// Where it stands: the open parenthesis or the call statement it is an argument of.
	let depth = 0;
	let open = -1;
	for (let k = at - 1; k >= 0; k--) {
		const raw = toks[k].rawText;
		if (raw === ')') {
			depth++;
		} else if (raw === '(') {
			if (depth === 0) {
				open = k;
				break;
			}
			depth--;
		}
	}
	const before = toks[at - 1];
	const after = toks[last + 1];
	// A call statement's argument: `Take AddressOf Cb`, `c.Add AddressOf Cb`.
	const calleeEnd = toks.findIndex((tok, k) => k > 0 && (k % 2 === 1 ? tok.rawText !== '.' : tokenName(tok) === undefined));
	const chainEnd = calleeEnd < 0 ? toks.length : calleeEnd;
	const statementArgument = open < 0 && at >= 1 && tokenName(toks[0]) !== undefined && !toks.some((tok, k) => k < at && tok.rawText === '=')
		&& (at === chainEnd || (before?.rawText === ',' && at > chainEnd));
	// `Debug.Print AddressOf Cb`: Print takes no AddressOf, whatever it names.
	if (open < 0 && tokenText(toks[0]) === 'debug' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'print') {
		return { why: `Debug.Print takes no 'AddressOf ${text}'.`, error: 'Syntax error' };
	}
	const startsSlot = before?.rawText === '(' || before?.rawText === ',' || statementArgument;
	if (!startsSlot || (open < 0 && !statementArgument)) {
		return { why: `AddressOf can stand only as an argument, and 'AddressOf ${text}' is not one.`, error: 'Syntax error' };
	}
	if (after && after.rawText !== ',' && after.rawText !== ')') {
		return { why: `'AddressOf ${text}' must be the whole argument, and an operator follows it.`, error: 'Argument not optional' };
	}
	const callee = open >= 0 ? toks[open - 1] : toks[0];
	const calleeName = tokenName(callee)?.toLowerCase();
	if (open >= 0 && (!calleeName || toks[open - 1]?.rawText === '(')) {
		return { why: `'AddressOf ${text}' cannot stand in parentheses of its own.`, error: 'Syntax error' };
	}
	const member = open >= 0 ? toks[open - 2]?.rawText === '.' : toks[1]?.rawText === '.';
	// Some VBA functions take it and some do not: VarPtr, StrPtr, Hex,
	// IsEmpty and TypeName compile; these do not (measured in Excel 16.0).
	const refused = calleeName && !member && !scope.signatures.has(calleeName) ? ADDRESSOF_REFUSED_BY.get(calleeName) : undefined;
	if (refused) {
		return { why: `${callee.rawText} takes no 'AddressOf ${text}'.`, error: refused };
	}
	// What it names.
	if (qualified) {
		const procedures = scope.standardModules.get(name.toLowerCase());
		const procedure = tokenName(toks[at + 3])!.toLowerCase();
		if (!procedures?.has(procedure)) {
			return scope.explicit ? { why: `'${text}' names no procedure of a standard module, and AddressOf takes only those.`, error: 'Variable not defined' } : undefined;
		}
	} else {
		const lower = name.toLowerCase();
		const kind = scope.moduleMembers.get(lower);
		if (scope.locals.has(lower) || kind === 'moduleVariable' || kind === 'constant') {
			return { why: `'${name}' is a variable, and AddressOf takes a Sub, Function or Property.`, error: 'Expected Sub, Function, or Property' };
		}
		if (kind === 'declare') {
			return { why: `'${name}' is a Declare, whose address AddressOf cannot take.`, error: 'Invalid use of AddressOf operator' };
		}
		const procedure = kind === 'sub' || kind === 'function' || kind?.startsWith('property') || scope.projectProcedureNames.has(lower);
		if (!procedure) {
			// Len, Array and the other reserved words are reserved-keyword-in-expression's.
			if (VALUE_WORD_ERRORS.has(lower)) {
				return undefined;
			}
			if (resolveRuntimeFunction(lower)) {
				return { why: `'${name}' is a VBA function, and AddressOf takes only the project's procedures.`, error: 'Invalid use of AddressOf operator' };
			}
			return scope.explicit ? { why: `'${name}' names no procedure of the project.`, error: 'Variable not defined' } : undefined;
		}
	}
	// AddressOf gives a LongPtr, which a ByVal Long takes no more than a LongLong (#298).
	const signature = calleeName && !member ? scope.signatures.get(calleeName) : undefined;
	if (signature && scope.is64) {
		let position = 0;
		for (let k = (open >= 0 ? open + 1 : 1); k < at; k++) {
			if (toks[k].rawText === ',') {
				position++;
			}
		}
		const param = signature.params[position];
		const type = normalizeType(param?.type);
		if (param && !param.byRef && type && NARROW.has(type)) {
			return { why: `'AddressOf ${text}' is a LongPtr in 64-bit VBA, and the ByVal ${param.type} '${param.name}' of '${signature.name}' takes no LongPtr.`, error: 'Type mismatch' };
		}
	}
	return undefined;
}
