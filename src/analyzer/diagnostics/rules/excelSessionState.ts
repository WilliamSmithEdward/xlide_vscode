// Workbook names are checked only for freshly added local sheets explicitly
// belonging to ThisWorkbook. ActiveWorkbook may change during NewSheet events.
// Clipboard contents and newly added sheet contents are runtime facts.
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { BodyNode, ModuleNode, ProcedureNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { walkEnteringBlocks } from '../dataflow';
import { stringLiteralValue } from '../typeInference';
import { activeModuleMembers, forEachVariableGroup, setAssignmentTarget, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';
import { namesIn } from './shared';

export function checkExcelSessionState(
    source: string,
    mod: ModuleNode,
    _callables: ReadonlySet<string>,
    memberCtx: MemberCompletionContext,
    activity: ConditionalActivityTracker | undefined,
    push: PushFn,
    procedureFilter?: (member: ProcedureNode) => boolean,
): void {
    const host = memberCtx.model?.hostName;
    if (host !== undefined && host !== 'Excel') { return; }
    const shadowsWorkbook = mod.members.some(node => node.kind === 'VariableGroup'
        ? node.declarations.some(decl => decl.name.toLowerCase() === 'thisworkbook')
        : 'name' in node && node.name?.toLowerCase() === 'thisworkbook');
    for (const member of activeModuleMembers(mod, activity)) {
        if (member.kind !== 'Procedure' || (procedureFilter && !procedureFilter(member))) { continue; }
        if (/\bon\s+error\b/i.test(source.slice(member.span.start, member.span.end))) { continue; }
        const locals = new Set<string>();
        forEachVariableGroup(member.body, group => {
            if (!group.isConst && group.modifier.toLowerCase() !== 'static') {
                for (const decl of group.declarations) { locals.add(decl.name.toLowerCase()); }
            }
        }, activity);
        if (shadowsWorkbook || locals.has('thisworkbook') || member.params.some(param => param.name.toLowerCase() === 'thisworkbook')) { continue; }
        const state = new Map<string, string>();
        const forget = (names: Iterable<string>): void => {
            for (const lower of names) { state.delete(lower); }
        };
        const visit = (node: BodyNode): void => {
            if (!isLeafStatement(node)) { return; }
            if (statementLabelDeclaration(source, node.span) || (node.kind === 'Statement' && node.singleLineIfBranches)) { state.clear(); }
            const toks = statementTokensAfterLeadingLabel(source, node.span);
            const set = setAssignmentTarget(source, node.span);
            if (set) {
                const lower = set.name.toLowerCase();
                const value = set.valueTokens.filter(tok => tok.kind !== 'comment').map(tok => tokenText(tok)).join('');
                if (locals.has(lower) && /^thisworkbook\.(?:worksheets|sheets)\.add(?:\(\))?$/.test(value)) {
                    // NewSheet handlers can rename existing sheets during Add.
                    // ThisWorkbook still establishes which workbook owns them.
                    for (const name of state.keys()) { state.set(name, ''); }
                    state.set(lower, '');
                } else { state.clear(); }
                return;
            }
            const sheet = toks.length === 5 && toks[1].rawText === '.' && tokenText(toks[2]) === 'name' && toks[3].rawText === '='
                ? tokenName(toks[0])?.toLowerCase() : undefined;
            if (sheet !== undefined && state.has(sheet) && toks[4].kind === 'stringLiteral') {
                const name = stringLiteralValue(toks[4].rawText);
                const taken = [...state].find(([other, held]) => other !== sheet && held !== '' && held.toLowerCase() === name.toLowerCase());
                if (taken) {
                    push('sheetNameInvalid', `"${name}" is the name the code gave '${taken[0]}', another sheet it added to ThisWorkbook, and sheet names ignore case. This will raise Run-time error '1004': That name is already taken.`, { start: node.span.start + toks[4].start, end: node.span.start + toks[4].end });
                }
                state.set(sheet, name);
                return;
            }
            state.clear();
        };
        walkEnteringBlocks(source, member.body, node => activity?.isInactive(node.span) === true, visit, {
            snapshot: () => new Map(state),
            restore: saved => { state.clear(); for (const [key, value] of saved) { state.set(key, value); } },
            forget,
            touches: stmt => new Set([...namesIn(source, stmt.span), ...state.keys()]),
        });
    }
}
