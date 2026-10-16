// Rule: member-exists-in-base-class.
//
// A document module, an Access form or report module, and a UserForm's
// module each derive from a host class: ThisWorkbook from Workbook, a sheet
// from Worksheet or Chart, ThisDocument from Document, Form_X from Form,
// Report_X from Report, a UserForm from UserForm. A module-level declaration
// may not take a name that class already has - the VBE refuses it with
// "Member already exists in an object module from which this object module
// derives". Measured in Excel, Word, PowerPoint and Access 16.0
// (2026-10-16):
//
//   - It clashes: a Sub, Function or Property Get, Let or Set, Public,
//     Private, Friend or with no modifier; a module-level variable, Dim,
//     Private or Public; a Private Const; a Private Declare. Any letter
//     case, and hidden members count (Document.Dummy1).
//   - It does not: a WithEvents variable, an Enum or its members, a Type, an
//     Event, a local, a Static local, a parameter, a label, or a name only
//     the Application has (Selection, CurrentDb).
//
// Not reported here, because the VBE refuses them for another reason first:
// a keyword such as Close ("Expected: identifier"), and a Public Const,
// Declare or array, which an object module may not make public at all.
//
// A sheet module whose sheet type is unknown is held only to the members a
// Worksheet and a Chart share. VB6's designer classes were not measured, so
// a VB6 module is left alone.

import { getExcelObjectModel } from '../../host/excelObjectModel';
import { isReservedIdentifier } from '../../lexer/keywordTable';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { ModuleNode, Span } from '../../parser/nodes';
import { designerClassMemberNames, ownObjectMemberNames, type AnalyzeModuleOptions, type PushFn } from '../analysisContext';
import { activeModuleMembers } from '../walker';

/** The Office designer classes this rule was measured against. */
const MEASURED_DESIGNER_CLASS = /^(Excel|Word|PowerPoint|Access)\./i;

interface BaseClass {
	/** What the message calls the class: `Word.Document`, `UserForm`. */
	label: string;
	/** Its members, lowercased. */
	members: ReadonlySet<string>;
}

/** The host class the module derives from, when it is known for certain. */
function baseClassOf(opts: AnalyzeModuleOptions): BaseClass | undefined {
	if (opts.designerClass) {
		return MEASURED_DESIGNER_CLASS.test(opts.designerClass)
			? { label: opts.designerClass, members: designerClassMemberNames(opts.designerClass, opts.hostModel ?? getExcelObjectModel()) }
			: undefined;
	}
	if (opts.moduleKind === 'userform') {
		return (opts.host ?? '').toLowerCase() === 'vb6' ? undefined : { label: 'UserForm', members: ownObjectMemberNames(opts) };
	}
	if (opts.moduleKind !== 'document') {
		return undefined;
	}
	const model = opts.hostModel ?? getExcelObjectModel();
	const host = (model.hostName ?? 'Excel').toLowerCase();
	const name = opts.moduleName?.toLowerCase();
	const known = opts.documentType === 'worksheet' ? 'Excel.Worksheet'
		: opts.documentType === 'chart' ? 'Excel.Chart'
			: opts.documentType === 'workbook' ? 'Excel.Workbook'
				: opts.documentType === 'document' ? 'Word.Document'
					: host === 'excel' && name === 'thisworkbook' ? 'Excel.Workbook'
						: host === 'word' && name === 'thisdocument' ? 'Word.Document' : undefined;
	if (known) {
		return { label: known, members: designerClassMemberNames(known, model) };
	}
	if (host !== 'excel') {
		return undefined;
	}
	// A sheet of either kind: only what both kinds have is certain to clash.
	let shared = SHEET_MEMBERS.get(model);
	if (!shared) {
		const chart = designerClassMemberNames('Excel.Chart', model);
		shared = new Set([...designerClassMemberNames('Excel.Worksheet', model)].filter((member) => chart.has(member)));
		SHEET_MEMBERS.set(model, shared);
	}
	return { label: 'Excel.Worksheet', members: shared };
}

/** What a Worksheet and a Chart share, per model: built once, read per analysis. */
const SHEET_MEMBERS = new WeakMap<object, ReadonlySet<string>>();

export function checkBaseClassMemberClash(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	opts: AnalyzeModuleOptions,
	push: PushFn,
): void {
	const base = baseClassOf(opts);
	if (!base || base.members.size === 0) {
		return;
	}
	const report = (name: string, span: Span | undefined): void => {
		if (!span || isReservedIdentifier(name) || !base.members.has(name.toLowerCase())) {
			return;
		}
		push(
			'memberExistsInBaseClass',
			`'${name}' is already a member of ${base.label}, which this module derives from. This is a VBE compile error: Member already exists in an object module from which this object module derives.`,
			span,
		);
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			report(member.name, member.nameSpan);
		} else if (member.kind === 'Declare') {
			if ((member.visibility ?? '').toLowerCase() !== 'public') {
				report(member.name, member.nameSpan);
			}
		} else if (member.kind === 'VariableGroup' && !member.withEvents) {
			const isPublic = ['public', 'global'].includes(member.modifier.toLowerCase());
			for (const decl of member.declarations) {
				if (isPublic && (member.isConst || decl.isArray || decl.fixedLength !== undefined)) {
					continue;
				}
				report(decl.name, decl.nameSpan);
			}
		}
	}
}
