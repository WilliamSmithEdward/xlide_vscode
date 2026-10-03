import {
	spelledAnnotation,
	type Annotation,
	type ModuleAnnotations,
	PROCEDURE_HEADER,
	VARIABLE_DECLARATION,
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

interface ModuleLine {
	text: string;
	next?: ModuleLine;
}

class ModuleText {
	private readonly first: ModuleLine;
	private readonly eol: string;
	private indexed = false;
	private readonly procedures = new Map<string, ModuleLine[]>();
	private readonly variables = new Map<string, ModuleLine>();
	private readonly continuationEnds = new WeakMap<ModuleLine, ModuleLine>();
	private readonly ownedAttributes = new WeakMap<ModuleLine, Map<string, ModuleLine>>();

	constructor(source: string) {
		this.eol = source.includes('\r\n') ? '\r\n' : '\n';
		const lines = source.replace(/\r\n/g, '\n').split('\n');
		let next: ModuleLine | undefined;
		for (let i = lines.length - 1; i >= 0; i--) { next = { text: lines[i], next }; }
		this.first = next!;
	}

	toString(): string {
		const lines: string[] = [];
		for (let line: ModuleLine | undefined = this.first; line; line = line.next) { lines.push(line.text); }
		return lines.join(this.eol);
	}

	/** The final header line: the preamble and its attributes. */
	private headerEnd(): ModuleLine | undefined {
		let end: ModuleLine | undefined;
		for (let line: ModuleLine | undefined = this.first; line; line = line.next) {
			if (HEADER_PREAMBLE.test(line.text)) { end = line; continue; }
			if (end) { break; }
		}
		return end;
	}

	private moduleAttribute(attribute: string): ModuleLine | undefined {
		const end = this.headerEnd();
		if (!end) { return undefined; }
		for (let line: ModuleLine | undefined = this.first; line; line = line.next) {
			const match = MODULE_ATTRIBUTE.exec(line.text);
			if (match && match[1].toLowerCase() === attribute.toLowerCase()) { return line; }
			if (line === end) { break; }
		}
		return undefined;
	}

	setModule(attribute: string, value: string, changes: AttributeChange[], skipped: string[], canInsert: boolean): void {
		const line = this.moduleAttribute(attribute);
		if (line) {
			const was = MODULE_ATTRIBUTE.exec(line.text)![2];
			if (was !== value) {
				line.text = 'Attribute '+attribute+' = '+value;
				changes.push({ target: 'module', attribute, from: was, to: value });
			}
			return;
		}
		if (!canInsert) {
			skipped.push(attribute+' is not an attribute this kind of module carries.');
			return;
		}
		const end = this.headerEnd();
		if (!end) {
			skipped.push('the module has no header to put '+attribute+' in.');
			return;
		}
		end.next = { text: 'Attribute '+attribute+' = '+value, next: end.next };
		changes.push({ target: 'module', attribute, to: value });
	}

	/** Line references survive inserted attributes; target lookup is built once. */
	private indexTargets(): void {
		if (this.indexed) { return; }
		this.indexed = true;
		let declarations = true;
		const end = this.headerEnd();
		for (let line = end ? end.next : this.first; line; line = line.next) {
			const header = PROCEDURE_HEADER.exec(line.text);
			if (header) {
				declarations = false;
				const key = header[1].toLowerCase();
				const list = this.procedures.get(key) ?? [];
				list.push(this.continuedHeaderEnd(line));
				this.procedures.set(key, list);
			} else if (declarations) {
				const variable = VARIABLE_DECLARATION.exec(line.text);
				if (variable && !this.variables.has(variable[1].toLowerCase())) { this.variables.set(variable[1].toLowerCase(), line); }
			}
		}
	}

	private continuedHeaderEnd(line: ModuleLine): ModuleLine {
		let last = line;
		const continued: ModuleLine[] = [];
		while (last.next && last.text.trimEnd().endsWith('_')) {
			const cached = this.continuationEnds.get(last);
			if (cached) { last = cached; break; }
			continued.push(last);
			last = last.next;
		}
		for (const one of continued) { this.continuationEnds.set(one, last); }
		return last;
	}

	setMember(annotation: Annotation, attribute: string, value: string, changes: AttributeChange[], skipped: string[]): void {
		this.indexTargets();
		const owner = annotation.target!;
		// A property's Get/Let/Set occurrence was counted before any insertions.
		const header = this.procedures.get(owner.toLowerCase())?.[annotation.targetOccurrence ?? 0];
		if (!header) {
			skipped.push("no procedure named '"+owner+"' was found for "+spelledAnnotation(annotation.kind)+'.');
			return;
		}
		this.setOwned(header, owner, attribute, value, changes);
	}

	setVariable(annotation: Annotation, value: string, changes: AttributeChange[], skipped: string[]): void {
		this.indexTargets();
		const owner = annotation.target!;
		const line = this.variables.get(owner.toLowerCase());
		if (!line) {
			skipped.push("no module-level variable named '"+owner+"' was found for '@VariableDescription.");
			return;
		}
		this.setOwned(line, owner, 'VB_VarDescription', value, changes);
	}

	private setOwned(after: ModuleLine, owner: string, attribute: string, value: string, changes: AttributeChange[]): void {
		let attributes = this.ownedAttributes.get(after);
		if (!attributes) {
			attributes = new Map();
			for (let line = after.next; line; line = line.next) {
				const match = OWNED_ATTRIBUTE.exec(line.text);
				if (!match) { break; }
				const key = match[1].toLowerCase()+'.'+match[2].toLowerCase();
				if (!attributes.has(key)) { attributes.set(key, line); }
			}
			this.ownedAttributes.set(after, attributes);
		}
		const key = owner.toLowerCase()+'.'+attribute.toLowerCase();
		const existing = attributes.get(key);
		const text = 'Attribute '+owner+'.'+attribute+' = '+value;
		if (existing) {
			const was = existing.text.slice(existing.text.indexOf('=') + 1).trim();
			if (was !== value) {
				existing.text = text;
				changes.push({ target: owner, attribute, from: was, to: value });
			}
			return;
		}
		const line = { text, next: after.next };
		after.next = line;
		attributes.set(key, line);
		changes.push({ target: owner, attribute, to: value });
	}
}
