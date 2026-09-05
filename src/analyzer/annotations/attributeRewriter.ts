import {
	isModuleAnnotation,
	spelledAnnotation,
	type Annotation,
	type ModuleAnnotations,
} from './attributeAnnotations';

/**
 * Writes the hidden attributes an annotation names into a module's own text.
 *
 * The text is edited in place and otherwise left alone: every line that is not
 * the attribute being set stays where it was, byte for byte. The module is
 * about to be written back over the developer's code, and anything touched by
 * accident would be a change to it. An attribute this does not manage is
 * carried through untouched, and one the annotations say nothing about is left
 * as it was - taking an attribute away is a separate, deliberate act.
 *
 * Parity with xlide_vbide's AttributeRewriter.
 */

export interface AttributeChange {
	/** The procedure or variable, or `module`. */
	target: string;
	attribute: string;
	/** The previous value, absent when the attribute was added. */
	from?: string;
	to: string;
}

export interface AttributeRewriteResult {
	text: string;
	/** Empty when the text is exactly as it came in. */
	changes: AttributeChange[];
	/** Annotations that named something the text has not got. */
	skipped: string[];
}

const PROCEDURE_HEADER =
	/^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(\p{L}[\p{L}\p{N}_]*)/iu;
const VARIABLE_DECLARATION =
	/^\s*(?:dim|private|public|global)\s+(?:withevents\s+)?(\p{L}[\p{L}\p{N}_]*)\b(?!\s*\()/iu;
const OWNED_ATTRIBUTE =
	/^\s*Attribute\s+(\p{L}[\p{L}\p{N}_]*)\.(VB_[A-Za-z_]+(?:\.VB_[A-Za-z_]+)?)\s*=/iu;
const MODULE_ATTRIBUTE = /^\s*Attribute\s+(VB_[A-Za-z_]+)\s*=\s*(.*?)\s*$/i;
/** The VERSION / BEGIN / END preamble a class module opens with. */
const HEADER_PREAMBLE = /^\s*(?:VERSION\s|BEGIN\b|END\b|MultiUse\s|\s*Attribute\s)/i;

/** A VBA string literal: quotes doubled, the whole thing quoted. */
export function vbaAttributeLiteral(text: string): string {
	return `"${text.replace(/"/g, '""')}"`;
}

/**
 * What the editor stores for an Excel macro hotkey: the letter, a literal
 * backslash-n, and 14.
 */
export function excelInvokeFuncFor(letter: string): string {
	return `${letter}\\n14`;
}

/**
 * Writes every attribute the annotations name into `source`, which is the
 * module's whole text, hidden header included.
 */
export function applyAttributeAnnotations(
	source: string,
	annotations: ModuleAnnotations,
): AttributeRewriteResult {
	const text = new ModuleText(source);
	const changes: AttributeChange[] = [];
	const skipped: string[] = [];

	for (const annotation of annotations.annotations) {
		switch (annotation.kind) {
			case 'ModuleDescription':
				text.setModule('VB_Description', vbaAttributeLiteral(annotation.argument ?? ''), changes, skipped, true);
				break;
			case 'PredeclaredId':
				text.setModule('VB_PredeclaredId', 'True', changes, skipped, false);
				break;
			case 'Exposed':
				text.setModule('VB_Exposed', 'True', changes, skipped, false);
				break;
			case 'Description':
				text.setMember(annotation, 'VB_Description', vbaAttributeLiteral(annotation.argument ?? ''), changes, skipped);
				break;
			case 'DefaultMember':
				text.setMember(annotation, 'VB_UserMemId', '0', changes, skipped);
				break;
			case 'Enumerator':
				text.setMember(annotation, 'VB_UserMemId', '-4', changes, skipped);
				break;
			case 'ExcelHotkey':
				text.setMember(
					annotation,
					'VB_ProcData.VB_Invoke_Func',
					vbaAttributeLiteral(excelInvokeFuncFor(annotation.argument ?? '')),
					changes,
					skipped,
				);
				break;
			case 'VariableDescription':
				text.setVariable(annotation, vbaAttributeLiteral(annotation.argument ?? ''), changes, skipped);
				break;
		}
	}

	return { text: text.toString(), changes, skipped };
}

class ModuleText {
	private readonly lines: string[];
	private readonly eol: string;

	constructor(source: string) {
		this.eol = source.includes('\r\n') ? '\r\n' : '\n';
		this.lines = source.replace(/\r\n/g, '\n').split('\n');
	}

	toString(): string {
		return this.lines.join(this.eol);
	}

	/** One past the last line of the header: the preamble and its attributes. */
	private headerEnd(): number {
		let end = 0;
		for (let at = 0; at < this.lines.length; at += 1) {
			if (HEADER_PREAMBLE.test(this.lines[at])) {
				end = at + 1;
				continue;
			}
			if (end > 0) {
				break;
			}
		}
		return end;
	}

	private moduleAttributeIndex(attribute: string): number {
		const end = this.headerEnd();
		for (let at = 0; at < end; at += 1) {
			const match = MODULE_ATTRIBUTE.exec(this.lines[at]);
			if (match && match[1].toLowerCase() === attribute.toLowerCase()) {
				return at;
			}
		}
		return -1;
	}

	setModule(
		attribute: string,
		value: string,
		changes: AttributeChange[],
		skipped: string[],
		canInsert: boolean,
	): void {
		const at = this.moduleAttributeIndex(attribute);
		if (at >= 0) {
			const was = MODULE_ATTRIBUTE.exec(this.lines[at])![2];
			if (was !== value) {
				this.lines[at] = `Attribute ${attribute} = ${value}`;
				changes.push({ target: 'module', attribute, from: was, to: value });
			}
			return;
		}
		if (!canInsert) {
			// A standard module's header has no VB_PredeclaredId or VB_Exposed
			// line, and the editor gives those meaning only on a class.
			skipped.push(`${attribute} is not an attribute this kind of module carries.`);
			return;
		}
		const end = this.headerEnd();
		if (end === 0) {
			skipped.push(`the module has no header to put ${attribute} in.`);
			return;
		}
		this.lines.splice(end, 0, `Attribute ${attribute} = ${value}`);
		changes.push({ target: 'module', attribute, to: value });
	}

	/** The last line of the nth header named `name`, or -1. */
	private headerIndex(name: string, occurrence: number): number {
		let seen = 0;
		for (let at = this.headerEnd(); at < this.lines.length; at += 1) {
			const match = PROCEDURE_HEADER.exec(this.lines[at]);
			if (!match || match[1].toLowerCase() !== name.toLowerCase()) {
				continue;
			}
			if (seen++ !== occurrence) {
				continue;
			}
			// The attributes follow the header's LAST line, which is the last
			// one ending in a continuation.
			let last = at;
			while (last < this.lines.length - 1 && this.lines[last].trimEnd().endsWith('_')) {
				last += 1;
			}
			return last;
		}
		return -1;
	}

	private variableIndex(name: string): number {
		const end = this.headerEnd();
		for (let at = end; at < this.lines.length; at += 1) {
			if (PROCEDURE_HEADER.test(this.lines[at])) {
				// Past the declarations section; a module-level variable is not
				// down here.
				return -1;
			}
			const match = VARIABLE_DECLARATION.exec(this.lines[at]);
			if (match && match[1].toLowerCase() === name.toLowerCase()) {
				return at;
			}
		}
		return -1;
	}

	setMember(
		annotation: Annotation,
		attribute: string,
		value: string,
		changes: AttributeChange[],
		skipped: string[],
	): void {
		const owner = annotation.target!;
		// The leg was counted when the annotation was read; deriving it from a
		// line number here would be read against text this has already edited.
		const header = this.headerIndex(owner, annotation.targetOccurrence ?? 0);
		if (header < 0) {
			skipped.push(
				`no procedure named '${owner}' was found for ${spelledAnnotation(annotation.kind)}.`,
			);
			return;
		}
		this.setOwned(header, owner, attribute, value, changes);
	}

	setVariable(
		annotation: Annotation,
		value: string,
		changes: AttributeChange[],
		skipped: string[],
	): void {
		const owner = annotation.target!;
		const at = this.variableIndex(owner);
		if (at < 0) {
			skipped.push(`no module-level variable named '${owner}' was found for '@VariableDescription.`);
			return;
		}
		this.setOwned(at, owner, 'VB_VarDescription', value, changes);
	}

	private setOwned(
		after: number,
		owner: string,
		attribute: string,
		value: string,
		changes: AttributeChange[],
	): void {
		const at = this.ownedIndex(after, owner, attribute);
		const line = `Attribute ${owner}.${attribute} = ${value}`;
		if (at >= 0) {
			const was = this.lines[at].slice(this.lines[at].indexOf('=') + 1).trim();
			if (was !== value) {
				this.lines[at] = line;
				changes.push({ target: owner, attribute, from: was, to: value });
			}
			return;
		}
		this.lines.splice(after + 1, 0, line);
		changes.push({ target: owner, attribute, to: value });
	}

	/** An existing `Attribute Owner.Name = ...` in the run below a header. */
	private ownedIndex(after: number, owner: string, attribute: string): number {
		for (let at = after + 1; at < this.lines.length; at += 1) {
			const match = OWNED_ATTRIBUTE.exec(this.lines[at]);
			if (!match) {
				// The attribute run ends at the first line that is not one.
				return -1;
			}
			if (match[1].toLowerCase() === owner.toLowerCase()
				&& match[2].toLowerCase() === attribute.toLowerCase()) {
				return at;
			}
		}
		return -1;
	}
}
