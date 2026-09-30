// Rule: an event handler's declaration must match its event (issue #195).
// Measured in Excel 16.0 (build 20326, 2026-09-29): a handler that differs
// from its event is refused with "Procedure declaration does not match
// description of event or procedure having the same name".
//
//  - Each parameter's passing: `app_SheetChange(Sh As Object, ...)` is
//    refused, since the event passes Sh ByVal. An event's ByRef parameter,
//    `Cancel As Boolean`, may be written ByRef or plain, not ByVal.
//  - Each parameter's type: Variant, a missing As, or Worksheet for the
//    event's Object is refused. A library prefix is the same type
//    (`Excel.Range`), and an enum parameter may be declared As Long.
//  - The count, and no Optional or ParamArray in place of the event's own.
//  - A Sub: a Function of the event's name is refused.
//
// Names differ freely, and Public or Private both compile. The VBE checks a
// handler only when its body holds something: an empty one compiles
// whatever its declaration.
//
// The handlers judged: `<variable>_<Event>` for a WithEvents variable of an
// Office or MSForms class, a document module's own object (Workbook_,
// Worksheet_, Chart_, Word's Document_), a UserForm's (UserForm_) and its
// controls'. The events come from the type libraries
// (host/eventSignaturesData.ts).
//
// A WithEvents variable of a project class takes the events the class
// declares (issue #220, measured in Excel 16.0). The same checks hold, an
// empty handler compiles here too, and a parameter of a project Enum may be
// declared As Long.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ModuleNode, ParameterNode, ProcedureNode } from '../../parser/nodes';
import { HOST_EVENT_SIGNATURES } from '../../host/eventSignaturesData';
import { resolveHostAlias } from '../../host/hostModel';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import type { AnalyzeModuleOptions, PushFn } from '../analysisContext';
import { normalizeType } from '../typeInference';
import { activeModuleMembers } from '../walker';

/** The events of each class, by lowercased qualified name, read once. */
const EVENTS_BY_CLASS = new Map<string, { name: string; events: Map<string, { name: string; params: string }> }>();
for (const [className, events] of Object.entries(HOST_EVENT_SIGNATURES)) {
	const byName = new Map<string, { name: string; params: string }>();
	for (const [event, params] of Object.entries(events)) {
		byName.set(event.toLowerCase(), { name: event, params });
	}
	EVENTS_BY_CLASS.set(className.toLowerCase(), { name: className, events: byName });
}

/** A document module's own object: the prefix its handlers take, and its class. */
const DOCUMENT_OBJECTS: Readonly<Record<string, readonly [string, string]>> = {
	workbook: ['Workbook', 'Excel.Workbook'],
	worksheet: ['Worksheet', 'Excel.Worksheet'],
	chart: ['Chart', 'Excel.Chart'],
	document: ['Document', 'Word.Document'],
};

export function checkEventHandlerSignatures(
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	opts: AnalyzeModuleOptions,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Each prefix a handler may take, with the classes whose events it raises.
	const sources = new Map<string, string[]>();
	const projectSources = new Map<string, { owner: string; events: Map<string, { name: string; params: string }> }>();
	const add = (prefix: string, ...classes: string[]): void => {
		const known = classes.filter((className) => EVENTS_BY_CLASS.has(className.toLowerCase()));
		if (known.length > 0) {
			sources.set(prefix.toLowerCase(), known);
		}
	};
	if (moduleKind === 'document' && opts.documentType && DOCUMENT_OBJECTS[opts.documentType]) {
		add(...DOCUMENT_OBJECTS[opts.documentType]);
	}
	if (moduleKind === 'userform') {
		add('UserForm', 'MSForms.UserForm');
		// A control's handlers take its own events and the extender's
		// (Enter, Exit, BeforeUpdate, AfterUpdate).
		for (const control of opts.implicitMembers ?? []) {
			add(control.name, control.type, 'MSForms.Control');
		}
	}
	const procedures: ProcedureNode[] = [];
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && member.withEvents) {
			for (const decl of member.declarations) {
				const project = projectEventsOf(decl.asType, memberCtx);
				if (project) {
					projectSources.set(decl.name.toLowerCase(), project);
					continue;
				}
				const className = eventClassOf(decl.asType, memberCtx);
				if (className) {
					add(decl.name, className);
				}
			}
		} else if (member.kind === 'Procedure') {
			procedures.push(member);
		}
	}
	if (sources.size === 0 && projectSources.size === 0) {
		return;
	}
	const projectEnums = new Set((memberCtx.projectClassMembers ?? []).filter((type) => type.kind === 'enum').map((type) => type.name.toLowerCase()));
	for (const proc of procedures) {
		// The VBE checks a handler only when its body holds a statement.
		if (proc.body.length === 0) {
			continue;
		}
		const underscore = proc.name.lastIndexOf('_');
		if (underscore <= 0) {
			continue;
		}
		const prefix = proc.name.slice(0, underscore).toLowerCase();
		const eventName = proc.name.slice(underscore + 1).toLowerCase();
		const projectSource = projectSources.get(prefix);
		const projectEvent = projectSource?.events.get(eventName);
		if (projectSource && projectEvent) {
			const problem = mismatch(proc, projectEvent.params, (type) => projectEnums.has(bareType(type)));
			if (problem) {
				push(
					'eventHandlerSignature',
					`'${proc.name}' does not match the event ${projectSource.owner}.${projectEvent.name}(${projectEvent.params}): ${problem}. This is a VBE compile error: Procedure declaration does not match description of event or procedure having the same name.`,
					proc.nameSpan ?? proc.span,
				);
			}
			continue;
		}
		const classes = sources.get(prefix);
		const found = classes
			?.map((className) => ({ owner: EVENTS_BY_CLASS.get(className.toLowerCase())!, event: EVENTS_BY_CLASS.get(className.toLowerCase())!.events.get(eventName) }))
			.find((entry) => entry.event !== undefined);
		if (!found?.event) {
			continue;
		}
		const problem = mismatch(proc, found.event.params, (type) => /^(xl|wd|pp|mso|fm)[a-z]/i.test(type));
		if (problem) {
			const owner = found.owner.name.slice(found.owner.name.indexOf('.') + 1);
			push(
				'eventHandlerSignature',
				`'${proc.name}' does not match the event ${owner}.${found.event.name}(${found.event.params}): ${problem}. This is a VBE compile error: Procedure declaration does not match description of event or procedure having the same name.`,
				proc.nameSpan ?? proc.span,
			);
		}
	}
}

/**
 * The events a project class declares, when a WithEvents variable's type
 * names one: the project's types come before the libraries', as in VBA.
 */
function projectEventsOf(
	asType: string | undefined,
	memberCtx: MemberCompletionContext,
): { owner: string; events: Map<string, { name: string; params: string }> } | undefined {
	const bare = asType?.trim().slice(asType.trim().lastIndexOf('.') + 1).toLowerCase();
	const type = bare ? (memberCtx.projectClassMembers ?? []).find((candidate) => candidate.name.toLowerCase() === bare) : undefined;
	if (!type) {
		return undefined;
	}
	const events = new Map<string, { name: string; params: string }>();
	for (const member of type.members) {
		const signature = member.signature ?? '';
		const open = signature.indexOf('(');
		if (member.kind === 'event' && open >= 0 && signature.endsWith(')')) {
			events.set(member.name.toLowerCase(), { name: member.name, params: signature.slice(open + 1, -1) });
		}
	}
	return { owner: type.name, events };
}

/** The class with events a WithEvents variable's declared type names, or undefined. */
function eventClassOf(asType: string | undefined, memberCtx: MemberCompletionContext): string | undefined {
	if (!asType) {
		return undefined;
	}
	const direct = EVENTS_BY_CLASS.get(asType.toLowerCase());
	if (direct) {
		return direct.name;
	}
	const resolved = resolveHostAlias(asType, memberCtx.model);
	return resolved && EVENTS_BY_CLASS.has(resolved.toLowerCase()) ? resolved : undefined;
}

interface EventParam {
	byVal: boolean;
	type: string;
	isArray: boolean;
}

function parseEventParams(params: string): EventParam[] {
	if (params.trim().length === 0) {
		return [];
	}
	return params.split(',').map((part) => {
		const text = part.trim();
		const typeMatch = /\sAs\s+(\S+)$/i.exec(text);
		const name = text.replace(/^(?:ByVal|ByRef)\s+/i, '').split(/\s/)[0];
		return { byVal: /^ByVal\s/i.test(text), type: typeMatch ? typeMatch[1] : 'Variant', isArray: name.endsWith('()') };
	});
}

/** Why the handler differs from the event, or undefined when it matches. */
function mismatch(proc: ProcedureNode, eventParams: string, isEnum: (type: string) => boolean): string | undefined {
	if (proc.procKind !== 'Sub') {
		return 'an event handler is a Sub';
	}
	const expected = parseEventParams(eventParams);
	const actual = proc.params;
	if (actual.length !== expected.length) {
		return `the event passes ${expected.length} parameter${expected.length === 1 ? '' : 's'}, this Sub takes ${actual.length}`;
	}
	for (let i = 0; i < actual.length; i++) {
		const param = actual[i];
		const want = expected[i];
		const label = `parameter ${i + 1}, '${param.name}'`;
		if (param.optional || param.paramArray) {
			return `${label} is ${param.optional ? 'Optional' : 'a ParamArray'}, and the event's is not`;
		}
		if (Boolean(param.isArray) !== want.isArray) {
			return param.isArray ? `${label} is an array, and the event's is not` : `${label} is not an array, and the event's is`;
		}
		if (param.byVal !== want.byVal) {
			return want.byVal ? `${label} must be ByVal` : `${label} must be ByRef, not ByVal`;
		}
		if (!sameType(param, want.type, isEnum)) {
			return `${label} is ${declaredType(param)}, and the event's is ${want.type}`;
		}
	}
	return undefined;
}

function declaredType(param: ParameterNode): string {
	return param.asType ?? (param.typeSuffix ? `${param.typeSuffix} (a type suffix)` : 'Variant');
}

/**
 * Whether a parameter's type is the event's. A library prefix names the same
 * type, and an enum - an Office or MSForms one (XlXmlExportResult, fmAction)
 * or the project's own - may be declared As Long, as measured.
 */
function sameType(param: ParameterNode, expected: string, isEnum: (type: string) => boolean): boolean {
	const suffixTypes: Readonly<Record<string, string>> = { '%': 'integer', '&': 'long', '!': 'single', '#': 'double', '@': 'currency', '$': 'string' };
	const actual = param.asType
		? bareType(param.asType)
		: param.typeSuffix ? suffixTypes[param.typeSuffix] ?? '' : 'variant';
	const want = bareType(expected);
	if (actual === want) {
		return true;
	}
	return actual === 'long' && isEnum(expected);
}

function bareType(type: string): string {
	const bare = type.trim().slice(type.trim().lastIndexOf('.') + 1);
	return normalizeType(bare) ?? bare.toLowerCase();
}
