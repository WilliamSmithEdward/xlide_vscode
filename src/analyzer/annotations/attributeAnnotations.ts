/**
 * The annotations that stand for a module's hidden attributes, Rubberduck's
 * convention.
 *
 * A VBA module carries attributes the code pane never shows - `VB_PredeclaredId`,
 * `VB_Description`, `VB_UserMemId` and the rest - and the editor offers no way
 * to set them. They decide whether a class has a default instance, what
 * `For Each` walks, and what IntelliSense says about a member. An annotation is
 * a comment naming the attribute the developer wants, written in the code where
 * it can be read, diffed and reviewed; the rewriter then writes the attribute to
 * match.
 *
 * PLACEMENT decides what an annotation binds to. A module annotation lives in
 * the declarations section, above the first procedure, anywhere in it. A member
 * annotation lives in the run of comment and blank lines directly above a
 * procedure's header and binds to that procedure; the same run above a
 * module-level variable binds a `'@VariableDescription` to the variable. A
 * member annotation above anything else, or a module annotation below the first
 * procedure, is a problem reported with its line rather than a guess.
 *
 * Read leniently, written one way: the name in any case, the argument in
 * brackets with or without quotes, or after a space.
 *
 * Parity with xlide_vbide's AttributeAnnotations.
 */

export type AnnotationKind =
	| 'ModuleDescription'
	| 'PredeclaredId'
	| 'Exposed'
	| 'Description'
	| 'DefaultMember'
	| 'Enumerator'
	| 'ExcelHotkey'
	| 'VariableDescription';

const KINDS: readonly AnnotationKind[] = [
	'ModuleDescription', 'PredeclaredId', 'Exposed', 'Description',
	'DefaultMember', 'Enumerator', 'ExcelHotkey', 'VariableDescription',
];

export interface Annotation {
	kind: AnnotationKind;
	/** 1-based line of the comment. */
	line: number;
	/** The text between the brackets, unquoted; absent for a bare annotation. */
	argument?: string;
	/** The procedure or variable it binds to; absent for a module annotation. */
	target?: string;
	/** 1-based line of that procedure's header or variable's declaration. */
	targetLine?: number;
	/**
	 * Which procedure of that name this is, counting from zero.
	 *
	 * A property's Get, Let and Set share a name, so the leg has to be named
	 * some other way. Counted HERE, while the line numbers are still the ones
	 * the developer wrote: the rewriter inserts attribute lines as it goes, and
	 * working the leg out from a line number afterwards puts the attribute on
	 * whichever leg the inserts have shifted past it.
	 */
	targetOccurrence?: number;
}

/** An annotation that cannot mean what it says where it sits, with the reason. */
export interface AnnotationProblem {
	line: number;
	message: string;
}

export interface ModuleAnnotations {
	annotations: Annotation[];
	problems: AnnotationProblem[];
}

// '@Name("arg")  '@Name(arg)  '@Name "arg"  '@Name arg  '@Name
const ANNOTATION_LINE =
	/^\s*'\s*@([A-Za-z]+)(?=\s|\(|"|$)\s*(?:\(\s*(?:"((?:[^"]|"")*)"|([^)]*?))\s*\)|"((?:[^"]|"")*)"|(\S+))?\s*$/;

const PROCEDURE_HEADER =
	/^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/**
 * A module-level variable. Const, Type, Enum, Declare, Event and Implements are
 * declarations too, but none of them carries a `VB_VarDescription`.
 */
const VARIABLE_DECLARATION =
	/^\s*(?:dim|private|public|global)\s+(?:withevents\s+)?(\p{L}[\p{L}\p{N}_]*)\b(?!\s*\()/iu;

const BLANK_OR_COMMENT = /^\s*(?:'|Rem\b|$)/i;

/** The documented spelling of a kind, apostrophe and at-sign included. */
export function spelledAnnotation(kind: AnnotationKind): string {
	return `'@${kind}`;
}

/** Whether the kind binds to the module rather than a procedure or variable. */
export function isModuleAnnotation(kind: AnnotationKind): boolean {
	return kind === 'ModuleDescription' || kind === 'PredeclaredId' || kind === 'Exposed';
}

/** Whether the kind carries text. */
export function annotationNeedsArgument(kind: AnnotationKind): boolean {
	return kind === 'ModuleDescription' || kind === 'Description'
		|| kind === 'ExcelHotkey' || kind === 'VariableDescription';
}

/** A hotkey is one letter; its case says whether Shift is held. */
export function isHotkeyArgument(argument: string | undefined): boolean {
	return argument !== undefined && argument.length === 1 && /\p{L}/u.test(argument);
}

/**
 * A comment opening with an at-sign, anywhere. Every annotation matches this,
 * so a module with none can be answered without looking at a single line - and
 * that is most modules, on a path that runs on every save.
 */
const ANY_ANNOTATION = /'[ \t]*@/;

/** The annotations of a module's code, with the problems found on the way. */
export function readAttributeAnnotations(source: string | undefined): ModuleAnnotations {
	if (!source || !ANY_ANNOTATION.test(source)) {
		return { annotations: [], problems: [] };
	}

	const annotations: Annotation[] = [];
	const problems: AnnotationProblem[] = [];
	let pending: Array<{ kind: AnnotationKind; line: number; argument?: string }> = [];
	let inDeclarations = true;
	let defaultMemberAt = 0;
	const seenProcedures = new Map<string, number>();

	const lines = source.split('\n');
	for (let at = 0; at < lines.length; at += 1) {
		const line = lines[at].replace(/\r$/, '');
		const number = at + 1;

		const match = ANNOTATION_LINE.exec(line);
		if (match) {
			const kind = kindOf(match[1]);
			if (!kind) {
				// Not one of ours: '@Folder, '@Ignore, '@TestMethod and any
				// prose starting with an at-sign belong to someone else.
				continue;
			}
			const argument = argumentOf(match);
			if (!isModuleAnnotation(kind)) {
				pending.push({ kind, line: number, ...(argument !== undefined ? { argument } : {}) });
				continue;
			}
			if (!inDeclarations) {
				problems.push({
					line: number,
					message: `${spelledAnnotation(kind)} is a module annotation and belongs in the declarations section, above the first procedure.`,
				});
				continue;
			}
			if (annotationNeedsArgument(kind) && !argument) {
				problems.push({ line: number, message: needsArgumentMessage(kind) });
				continue;
			}
			if (annotations.some((one) => one.kind === kind)) {
				problems.push({
					line: number,
					message: `${spelledAnnotation(kind)} appears more than once; the first one counts.`,
				});
				continue;
			}
			annotations.push({ kind, line: number, ...(argument !== undefined ? { argument } : {}) });
			continue;
		}

		if (BLANK_OR_COMMENT.test(line)) {
			continue;
		}

		// A line of code. Whatever is pending binds to it, or fails here.
		const header = PROCEDURE_HEADER.exec(line);
		if (header) {
			inDeclarations = false;
			const lower = header[1].toLowerCase();
			const occurrence = seenProcedures.get(lower) ?? 0;
			seenProcedures.set(lower, occurrence + 1);
			bindToProcedure(header[1], number, occurrence);
			continue;
		}
		const variable = inDeclarations ? VARIABLE_DECLARATION.exec(line) : null;
		if (variable) {
			bindToVariable(variable[1], number);
			continue;
		}
		for (const one of pending) {
			problems.push({
				line: one.line,
				message: `${spelledAnnotation(one.kind)} is above a line that is not a `
					+ `${one.kind === 'VariableDescription' ? 'module-level variable' : 'procedure'}, `
					+ 'so there is nothing to bind it to.',
			});
		}
		pending = [];
	}

	for (const one of pending) {
		problems.push({
			line: one.line,
			message: `${spelledAnnotation(one.kind)} is above nothing, so there is nothing to bind it to.`,
		});
	}

	return { annotations, problems };

	function bindToProcedure(name: string, headerLine: number, occurrence: number): void {
		for (const one of pending) {
			if (one.kind === 'VariableDescription') {
				problems.push({
					line: one.line,
					message: `${spelledAnnotation(one.kind)} describes a module-level variable, and '${name}' is a procedure. Use '@Description for a procedure.`,
				});
				continue;
			}
			if (annotationNeedsArgument(one.kind) && !one.argument) {
				problems.push({ line: one.line, message: needsArgumentMessage(one.kind) });
				continue;
			}
			if (one.kind === 'ExcelHotkey' && !isHotkeyArgument(one.argument)) {
				problems.push({
					line: one.line,
					message: `${spelledAnnotation(one.kind)} takes one letter: a lower-case letter is Ctrl+letter, an upper-case one Ctrl+Shift+letter.`,
				});
				continue;
			}
			if (one.kind === 'DefaultMember') {
				if (defaultMemberAt !== 0) {
					problems.push({
						line: one.line,
						message: `'@DefaultMember appears again; a class has one default member, and line ${defaultMemberAt} already names it.`,
					});
					continue;
				}
				defaultMemberAt = one.line;
			}
			if (annotations.some((found) => found.kind === one.kind
				&& found.target === name && found.targetLine === headerLine)) {
				problems.push({
					line: one.line,
					message: `${spelledAnnotation(one.kind)} appears more than once above '${name}'; the first one counts.`,
				});
				continue;
			}
			annotations.push({
				kind: one.kind,
				line: one.line,
				...(one.argument !== undefined ? { argument: one.argument } : {}),
				target: name,
				targetLine: headerLine,
				targetOccurrence: occurrence,
			});
		}
		pending = [];
	}

	function bindToVariable(name: string, declarationLine: number): void {
		for (const one of pending) {
			if (one.kind !== 'VariableDescription') {
				problems.push({
					line: one.line,
					message: `${spelledAnnotation(one.kind)} describes a procedure, and '${name}' is a variable. Use '@VariableDescription for a variable.`,
				});
				continue;
			}
			if (!one.argument) {
				problems.push({ line: one.line, message: needsArgumentMessage(one.kind) });
				continue;
			}
			if (annotations.some((found) => found.kind === one.kind && found.target === name)) {
				problems.push({
					line: one.line,
					message: `${spelledAnnotation(one.kind)} appears more than once above '${name}'; the first one counts.`,
				});
				continue;
			}
			annotations.push({
				kind: one.kind,
				line: one.line,
				argument: one.argument,
				target: name,
				targetLine: declarationLine,
			});
		}
		pending = [];
	}
}

function needsArgumentMessage(kind: AnnotationKind): string {
	const spelled = spelledAnnotation(kind);
	return `${spelled} needs the text to write, in brackets: ${spelled}("...").`;
}

function kindOf(name: string): AnnotationKind | undefined {
	const lower = name.toLowerCase();
	return KINDS.find((kind) => kind.toLowerCase() === lower);
}

function argumentOf(match: RegExpExecArray): string | undefined {
	const [, , quoted, unquoted, bare, word] = match;
	if (quoted !== undefined) { return quoted.replace(/""/g, '"'); }
	if (bare !== undefined) { return bare.replace(/""/g, '"'); }
	if (unquoted !== undefined && unquoted.length > 0) { return unquoted; }
	if (word !== undefined) { return word; }
	return undefined;
}
