import {assignmentTargetFromTokens, assignmentTargetName} from '../completion/assignmentTarget';
import type {MemberCompletionContext} from '../completion/memberAccess';
import {resolveExactMemberCompletion} from './typeInference';
import {resolveBareIdentifierBinding} from '../symbols/nameResolution';
import {procedureParamsFromSymbol, type ModuleSymbols, type VbaSymbol, type VbaProcedureParam, type VbaProjectClassMembers} from '../symbols/symbolModel';
import {tokenName, tokensWithoutLeadingLineNumber} from '../lexer/tokenHelpers';
import type {Span} from '../parser/nodes';
import {statementTokensAfterLeadingLabel, topLevelOperatorIndex} from './walker';
import {validateArity, splitArgSlots} from './callExtraction';
import type {PushFn} from './analysisContext';

interface SetterNames {all: Set<string>; indexed: Set<string>}
const SYMBOL_SETTERS = new WeakMap<object, SetterNames>();
const SURFACE_SETTERS = new WeakMap<object, SetterNames>();
const MEMBER_SETTERS = new WeakMap<VbaProjectClassMembers, SetterNames>();
function surfaceSetterNames(surface: VbaProjectClassMembers): SetterNames {
    let names = MEMBER_SETTERS.get(surface);
    if (!names) {
        names = {all:new Set(), indexed:new Set()};
        for (const member of surface.members) {
            const params = [member.procedureParams?.propertyLet, member.procedureParams?.propertySet].filter(Boolean);
            if (!params.length) { continue; }
            const name = member.name.toLowerCase(); names.all.add(name);
            if (params.some(params => params!.length > 1)) { names.indexed.add(name); }
        }
        MEMBER_SETTERS.set(surface,names);
    }
    return names;
}
function symbolSetterNames(key: object, symbols: readonly VbaSymbol[]): SetterNames {
    let names = SYMBOL_SETTERS.get(key);
    if (!names) {
        names = {all:new Set(), indexed:new Set()};
        for (const symbol of symbols) {
            if (symbol.kind !== 'propertyLet' && symbol.kind !== 'propertySet') { continue; }
            const name = symbol.name.toLowerCase(); names.all.add(name);
            if ((symbol.children?.filter(child => child.kind === 'parameter').length ?? 0) > 1) { names.indexed.add(name); }
        }
        SYMBOL_SETTERS.set(key,names);
    }
    return names;
}
function mayBeSetter(symbols: ModuleSymbols, visible: readonly VbaSymbol[] | undefined, ctx: MemberCompletionContext, name: string, indexed: boolean): boolean {
    const has = (names: SetterNames) => (indexed ? names.all : names.indexed).has(name);
    if (has(symbolSetterNames(symbols.root, symbols.root.children ?? [])) || (visible && has(symbolSetterNames(visible,visible)))) { return true; }
    const surfaces = ctx.projectClassMembers;
    if (!surfaces) { return false; }
    let names = SURFACE_SETTERS.get(surfaces);
    if (!names) {
        names = {all:new Set(), indexed:new Set()};
        for (const surface of surfaces) {
            const contribution = surfaceSetterNames(surface);
            for (const name of contribution.all) { names.all.add(name); }
            for (const name of contribution.indexed) { names.indexed.add(name); }
        }
        SURFACE_SETTERS.set(surfaces,names);
    }
    return has(names);
}

export function sourceSetterAssignment(source: string, span: Span, symbols: ModuleSymbols, procedure: VbaSymbol | undefined,
    visible: readonly VbaSymbol[] | undefined, ctx: MemberCompletionContext) {
    let tokens = tokensWithoutLeadingLineNumber(statementTokensAfterLeadingLabel(source, span));
    if (tokens[0]?.rawText.toLowerCase() === 'if') { return undefined; } // Branches are visited separately.
    const usesSet = tokens[0]?.rawText.toLowerCase() === 'set';
    if (usesSet) { tokens = tokens.slice(1); }
    const equals = topLevelOperatorIndex(tokens, '=');
    if (equals < 0) { return undefined; }
    const candidate = assignmentTargetName(tokens.slice(0,equals));
    const candidateName = candidate && tokenName(tokens[candidate.index])?.toLowerCase();
    if (!candidate || !candidateName || !mayBeSetter(symbols,visible,ctx,candidateName,candidate.indexed)) { return undefined; }
    const target = assignmentTargetFromTokens(tokens.slice(0, equals + 1));
    const named = target && assignmentTargetName(target);
    if (!target || !named) { return undefined; }
    const token = target[named.index], name = tokenName(token);
    if (!name) { return undefined; }
    const kind = usesSet ? 'propertySet' : 'propertyLet';
    let params: readonly VbaProcedureParam[] | undefined;
    let legacyNoIndex = false;
    let qualifier: string | undefined;
    if (named.index === 0) {
        const binding = resolveBareIdentifierBinding({currentModule: symbols, enclosingProcedure: procedure, projectVisibleSymbols: visible, name, context: 'assignmentTarget'});
        if (binding.scope === 'ambiguous') { return undefined; }
        const setter = binding.definitions.find(def => def.kind === kind);
        if (!setter) { return undefined; }
        params = procedureParamsFromSymbol(setter);
        qualifier = setter.moduleName;
    } else {
        const member = resolveExactMemberCompletion(source, name, span.start + token.end, ctx);
        params = member?.procedureParams?.[kind];
        qualifier = member?.owner;
        legacyNoIndex = !usesSet && member?.signature === undefined && params?.length === 1;
    }
    if (!params?.length) { return undefined; }
    const indexParams = params.slice(0, -1).map(param => ({...param, optional: Boolean(param.optional), paramArray: Boolean(param.paramArray)}));
    const split = !named.indexed || target.length === named.index + 3 ? {slots:[], spans:[]} : splitArgSlots(target.slice(named.index + 2, -1), span.start);
    return {name, qualifier, indexed: named.indexed, nameSpan: {start: span.start + token.start, end: span.start + token.end}, indexParams, slots: split.slots, slotSpans: split.spans, sliceStart: span.start, legacyNoIndex};
}

export function invalidSetterAssignmentArity(assignment: NonNullable<ReturnType<typeof sourceSetterAssignment>>, source: string, push: PushFn): boolean {
    // propertyUse retains its established Invalid use of property diagnostic
    // for indexing a setter-only property whose Let has no index parameters.
    if (assignment.legacyNoIndex && assignment.slots.length) { return true; }
    let invalid = false;
    validateArity(source, {name: assignment.name, params: assignment.indexParams}, assignment, (rule, message, span, data) => {
        invalid = true;
        const placeholder = data?.missingRequiredArgumentPlaceholder;
        if (placeholder && !assignment.indexed && assignment.slots.length === 0) {
            data = {...data, missingRequiredArgumentPlaceholder: {...placeholder, edit: {...placeholder.edit, newText: `(${placeholder.edit.newText.trim()})`}}};
        }
        push(rule, assignment.slots.length === 0 && assignment.indexParams.some(param => !param.optional && !param.paramArray) ? `Argument not optional: property '${assignment.name}' requires an index.` : message, span, data);
    });
    return invalid;
}
