// Doc comments that do not match their declaration.
//
// A `'''` block written in XML - one holding any tag of the vocabulary in
// user_guides/vba-doc-comments.md - documents the Sub, Function, Property,
// Declare or Event below it, and hovers and call tips show it as the truth
// about that declaration. So once a procedure has one, it has to describe
// the whole surface a caller sees:
//
//   - every parameter, by a <param> naming it, once, with something in it;
//   - the value a Function returns, by a <returns>, which a Sub, a Property
//     Let or Set and an Event must not have, since they return nothing;
//   - every tag closed, and no tag but <param> given twice.
//
// A property is its value, and its <summary> describes it: neither the value
// a Property Get returns nor the one a Property Let or Set receives in its
// last parameter needs a tag of its own, though either may have one. A block
// of plain text is a note, not XML, and is left alone, and so is a row of
// apostrophes drawn above a procedure.
//
// A finding about a tag points at the tag; one about something the block
// leaves out points at the declaration.

import type {
	DeclareNode,
	EventNode,
	ModuleNode,
	ParameterNode,
	ProcedureNode,
	Span,
} from '../../parser/nodes';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { leadingDocLines, scanDocTags } from '../../docs/docComment';
import type { DocBlockLine, DocTagOccurrence } from '../../docs/docComment';
import type { PushFn, VbaDocCommentFix } from '../analysisContext';
import { activeModuleMembers } from '../walker';
import { detectEol, leadingWhitespace, lineStartAt } from '../../../vbaSourceScan';

type DocumentedMember = ProcedureNode | DeclareNode | EventNode;

interface TextEdit {
	span: Span;
	newText: string;
}

/** What a declaration shows its callers, which its doc comment describes. */
interface Surface {
	/** `Sub`, `Function`, `Property Get`, `Property Let`, `Property Set` or `Event`. */
	label: string;
	params: readonly ParameterNode[];
	/** A Property Let or Set's last parameter: the property's value. */
	valueParam?: ParameterNode;
	returns: 'required' | 'optional' | 'none';
}

const PROCEDURE_LABELS: Record<ProcedureNode['procKind'], string> = {
	Sub: 'Sub',
	Function: 'Function',
	PropertyGet: 'Property Get',
	PropertyLet: 'Property Let',
	PropertySet: 'Property Set',
};

/** A `disable-next-line` directive, which is about the line below it only. */
const NEXT_LINE_DIRECTIVE_RE = /^\s*'+\s*@xlide-analysis-disable-next-line\b/i;

/** Tags a doc comment has at most one of; the parser reads the first. */
const SINGLE_TAGS = new Set(['summary', 'returns', 'remarks', 'example', 'signature']);

/**
 * Rule: a procedure's XML doc comment describes every parameter and return
 * value it has, and nothing it does not have.
 */
export function checkDocComments(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' && member.kind !== 'Declare' && member.kind !== 'Event') {
			continue;
		}
		const lines = leadingDocLines(source, member.span.start);
		const tags = lines.length > 0 ? scanDocTags(lines) : undefined;
		if (tags) {
			checkMember(new DocBlock(source, lines, tags), member, surfaceOf(member), push);
		}
	}
}

function surfaceOf(member: DocumentedMember): Surface {
	const params = member.params.filter((param) => param.name);
	switch (member.kind) {
		case 'Procedure': {
			const setter = member.procKind === 'PropertyLet' || member.procKind === 'PropertySet';
			const last = member.params[member.params.length - 1];
			return {
				label: PROCEDURE_LABELS[member.procKind],
				params,
				valueParam: setter && last?.name ? last : undefined,
				returns: member.procKind === 'Function'
					? 'required'
					: member.procKind === 'PropertyGet' ? 'optional' : 'none',
			};
		}
		case 'Declare':
			return {
				label: member.isFunction ? 'Function' : 'Sub',
				params,
				returns: member.isFunction ? 'required' : 'none',
			};
		case 'Event':
			return { label: 'Event', params, returns: 'none' };
	}
}

function checkMember(
	block: DocBlock,
	member: DocumentedMember,
	surface: Surface,
	push: PushFn,
): void {
	const report = (
		rule: Parameters<PushFn>[0],
		message: string,
		span: Span,
		fixes: VbaDocCommentFix[] = [],
	): void => push(rule, message, span, fixes.length > 0 ? { docCommentFixes: fixes } : undefined);

	// The tag each parameter is described by: the first naming it, as the
	// call tip reads it.
	const paramTags = new Map<string, DocTagOccurrence>();
	for (const tag of block.tags) {
		const lower = tag.name?.toLowerCase();
		if (tag.tag === 'param' && lower && !paramTags.has(lower)) {
			paramTags.set(lower, tag);
		}
	}
	const byName = new Map(surface.params.map((param) => [param.name.toLowerCase(), param]));
	const undescribed = surface.params.filter((param) => !paramTags.has(param.name.toLowerCase()));

	const seen = new Set<string>();
	let returns: DocTagOccurrence | undefined;
	for (const tag of block.tags) {
		const first = tag.tag === 'param'
			? !tag.name || paramTags.get(tag.name.toLowerCase()) === tag
			: !seen.has(tag.tag);
		seen.add(tag.tag);
		if (tag.tag === 'returns') {
			returns ??= tag;
		}
		if (tag.end === undefined) {
			report('docTagUnclosed', `This <${tag.tag}> is not closed; end it with </${tag.tag}>.`, tag.open);
			continue;
		}
		const empty = tag.text === '' && !tag.hasHints;
		if (!first && (tag.tag === 'param' || SINGLE_TAGS.has(tag.tag))) {
			report(
				'docTagDuplicate',
				tag.tag === 'param'
					? `The doc comment already describes parameter '${tag.name}'.`
					: `The doc comment already has a <${tag.tag}>, and only the first is shown.`,
				tag.nameSpan ?? tag.open,
				[removeFix(block, tag, `Remove the repeated <${tag.tag}>`)],
			);
			continue;
		}
		if (tag.tag === 'param') {
			if (!tag.name || !tag.nameSpan) {
				report(
					'docParamUnknown',
					'This <param> has no name="..." to say which parameter it describes.',
					tag.open,
					[
						...undescribed.map((param): VbaDocCommentFix => ({
							title: `Name the <param> '${param.name}'`,
							isPreferred: undescribed.length === 1,
							edits: [{ span: at(tag.open.start + '<param'.length), newText: ` name="${param.name}"` }],
						})),
						removeFix(block, tag, 'Remove the <param>'),
					],
				);
			} else if (!byName.has(tag.name.toLowerCase())) {
				const nameSpan = tag.nameSpan;
				report(
					'docParamUnknown',
					`'${member.name}' has no parameter named '${tag.name}'.`,
					nameSpan,
					[
						...undescribed.map((param): VbaDocCommentFix => ({
							title: `Rename the <param> to '${param.name}'`,
							isPreferred: undescribed.length === 1,
							edits: [{ span: nameSpan, newText: param.name }],
						})),
						removeFix(block, tag, `Remove the <param> for '${tag.name}'`),
					],
				);
			} else if (empty) {
				report('docParamMissing', `The <param> for '${byName.get(tag.name.toLowerCase())!.name}' is empty.`, tag.open);
			}
		} else if (tag.tag === 'returns') {
			if (surface.returns === 'none') {
				report(
					'docReturnsUnexpected',
					`${surface.label} '${member.name}' returns no value, but its doc comment describes one.`,
					tag.open,
					[removeFix(block, tag, 'Remove the <returns>')],
				);
			} else if (empty) {
				report('docReturnsMissing', 'The <returns> is empty.', tag.open);
			}
		}
	}

	const missing = undescribed.filter((param) => param !== surface.valueParam);
	const addAll: VbaDocCommentFix | undefined = missing.length > 1
		? {
			title: `Add the ${missing.length} missing <param> tags`,
			edits: mergeInsertions(missing.map((param) => addParamEdit(block, surface, param, paramTags))),
		}
		: undefined;
	for (const param of missing) {
		report(
			'docParamMissing',
			`The doc comment does not describe parameter '${param.name}'.`,
			param.nameSpan ?? param.span,
			[
				{
					title: `Add a <param> for '${param.name}'`,
					isPreferred: true,
					edits: [addParamEdit(block, surface, param, paramTags)],
				},
				...(addAll ? [addAll] : []),
			],
		);
	}
	if (surface.returns === 'required' && !returns) {
		report(
			'docReturnsMissing',
			`The doc comment does not describe what '${member.name}' returns.`,
			member.nameSpan ?? member.span,
			[{ title: 'Add a <returns>', isPreferred: true, edits: [addReturnsEdit(block)] }],
		);
	}
}

function at(offset: number): Span {
	return { start: offset, end: offset };
}

function removeFix(block: DocBlock, tag: DocTagOccurrence, title: string): VbaDocCommentFix {
	return { title, edits: [block.removeTag(tag)] };
}

/**
 * A new `<param>` goes among the others in signature order; with none to go
 * by, after the summary, or before whatever follows the parameters.
 */
function addParamEdit(
	block: DocBlock,
	surface: Surface,
	param: ParameterNode,
	paramTags: ReadonlyMap<string, DocTagOccurrence>,
): TextEdit {
	const content = `<param name="${param.name}"></param>`;
	const index = surface.params.indexOf(param);
	let beforeEnd: number | undefined;
	let afterStart: number | undefined;
	for (let i = 0; i < surface.params.length; i += 1) {
		const tag = paramTags.get(surface.params[i].name.toLowerCase());
		if (tag?.end === undefined) {
			continue;
		}
		if (i < index) {
			beforeEnd = tag.end;
		} else if (i > index) {
			afterStart ??= tag.open.start;
		}
	}
	if (beforeEnd !== undefined) {
		return block.insertLineAfter(beforeEnd, content);
	}
	if (afterStart !== undefined) {
		return block.insertLineBefore(afterStart, content);
	}
	return block.insertBeforeTrailingTags(content);
}

/** A new `<returns>` goes after the last `<param>`, or where one would go. */
function addReturnsEdit(block: DocBlock): TextEdit {
	const params = block.tags.filter((tag) => tag.tag === 'param' && tag.end !== undefined);
	const last = params[params.length - 1];
	return last?.end !== undefined
		? block.insertLineAfter(last.end, '<returns></returns>')
		: block.insertBeforeTrailingTags('<returns></returns>');
}

/** Insertions at one offset become one edit, in the order given. */
function mergeInsertions(edits: readonly TextEdit[]): TextEdit[] {
	const byOffset = new Map<number, TextEdit>();
	for (const edit of edits) {
		const merged = byOffset.get(edit.span.start);
		byOffset.set(edit.span.start, merged ? { span: merged.span, newText: merged.newText + edit.newText } : edit);
	}
	return [...byOffset.values()].sort((a, b) => a.span.start - b.span.start);
}

/** The `'''` block above one declaration, and the edits that change it. */
class DocBlock {
	private readonly eol: string;

	constructor(
		private readonly source: string,
		readonly lines: readonly DocBlockLine[],
		readonly tags: readonly DocTagOccurrence[],
	) {
		this.eol = detectEol(source);
	}

	insertLineBefore(offset: number, content: string): TextEdit {
		const line = this.lineAt(offset);
		return { span: at(line.directivesStart), newText: this.newLine(line, content) };
	}

	insertLineAfter(offset: number, content: string): TextEdit {
		const line = this.lineAt(offset);
		return { span: at(this.nextLineStart(line)), newText: this.newLine(line, content) };
	}

	/** After the summary, else before the returns, remarks or example, else last. */
	insertBeforeTrailingTags(content: string): TextEdit {
		const summary = this.tags.find((tag) => tag.tag === 'summary');
		if (summary?.end !== undefined) {
			return this.insertLineAfter(summary.end, content);
		}
		const trailing = this.tags.find((tag) => tag.tag === 'returns' || tag.tag === 'remarks' || tag.tag === 'example');
		if (trailing) {
			return this.insertLineBefore(trailing.open.start, content);
		}
		const written = this.lines.filter((line) => line.text.trim() !== '');
		return this.insertLineAfter(written[written.length - 1].start, content);
	}

	/**
	 * The lines a tag has to itself go with it, and so does a
	 * `disable-next-line` right above them, which would otherwise move on to
	 * the line after. Otherwise just the tag goes.
	 */
	removeTag(tag: DocTagOccurrence): TextEdit {
		const end = tag.end ?? tag.open.end;
		const first = this.lineAt(tag.open.start);
		const last = this.lineAt(end);
		const before = this.source.slice(first.textStart, tag.open.start);
		const after = this.source.slice(end, this.lineEnd(last));
		if (before.trim() !== '' || after.trim() !== '') {
			return { span: { start: tag.open.start, end }, newText: '' };
		}
		let start = first.start;
		if (first.directivesStart < first.start) {
			const above = lineStartAt(this.source, first.start - 1);
			if (NEXT_LINE_DIRECTIVE_RE.test(this.source.slice(above, first.start))) {
				start = above;
			}
		}
		return { span: { start, end: this.nextLineStart(last) }, newText: '' };
	}

	/** A `'''` line indented the way the line it goes next to is. */
	private newLine(beside: DocBlockLine, content: string): string {
		const prefix = this.source.slice(beside.start, beside.textStart) + leadingWhitespace(beside.text);
		return `${prefix}${content}${this.eol}`;
	}

	private lineAt(offset: number): DocBlockLine {
		let found = this.lines[0];
		for (const line of this.lines) {
			if (line.start <= offset) {
				found = line;
			}
		}
		return found;
	}

	private lineEnd(line: DocBlockLine): number {
		return line.textStart + line.text.length;
	}

	private nextLineStart(line: DocBlockLine): number {
		const lf = this.source.indexOf('\n', this.lineEnd(line));
		return lf < 0 ? this.source.length : lf + 1;
	}
}
