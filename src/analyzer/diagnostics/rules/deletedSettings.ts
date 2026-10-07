// Rule: DeleteSetting of a key, section or application the procedure
// already deleted (issue #700, measured in Excel 16.0).
//
// `DeleteSetting "App", "S", "k"` a second time raises 5, Invalid procedure
// call or argument, and so does deleting a key of a section or application
// deleted above, or that section or application again. A SaveSetting of the
// application in between, or a call, a label or a block that may save or
// delete settings, ends what is known. Under On Error Resume Next nothing is
// reported, and the setting is gone either way.
//
// Deleting a key the code never saved raises 5 too, but the registry keeps
// settings from earlier runs, so that is not for a static rule.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode , ProcedureNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { stringLiteralValue } from '../typeInference';
import { activeModuleMembers, isInactiveNode, matchParenFrom, statementTokensAfterLeadingLabel, tokenText } from '../walker';

/** The literal arguments of a SaveSetting or DeleteSetting, lowercased; undefined for one that is not a literal. */
function settingArguments(toks: readonly VbaToken[]): (string | undefined)[] {
	const open = toks[1]?.rawText === '(' ? 1 : -1;
	const end = open < 0 ? toks.length : matchParenFrom(toks, open);
	if (end < 0) {
		return [undefined];
	}
	return splitTopLevelTokenGroups(toks, open < 0 ? 1 : 2, ',', end).map((arg) => {
		const value = arg.filter((tok) => tok.kind !== 'comment');
		return value.length === 1 && value[0].kind === 'stringLiteral' ? stringLiteralValue(value[0].rawText).toLowerCase() : undefined;
	});
}

export function checkDeletedSettings(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	procedureFilter?: (member: ProcedureNode) => boolean,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && procedureFilter && !procedureFilter(member)) { continue; }
		if (member.kind !== 'Procedure' || !/\bdeletesetting\b/i.test(source.slice(member.span.start, member.span.end))) {
			continue;
		}
		// What the procedure deleted, as `app|section|key`, with `*` for all.
		const gone = new Set<string>();
		let resumeNext = false;
		const run = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (isInactiveNode(activity, node)) {
					continue;
				}
				if (!isLeafStatement(node)) {
					// A block may save or delete settings on one path only.
					gone.clear();
					if ('body' in node && Array.isArray(node.body)) {
						run(node.body as BodyNode[]);
					}
					gone.clear();
					continue;
				}
				if (statementLabelDeclaration(source, node.span)) {
					gone.clear();
				}
				const toks = statementTokensAfterLeadingLabel(source, node.span);
				const head = tokenText(toks[0]);
				if (head === 'on' && tokenText(toks[1]) === 'error') {
					resumeNext = tokenText(toks[2]) === 'resume';
					continue;
				}
				if (head === 'savesetting' && toks[1]?.rawText !== '=') {
					const [app] = settingArguments(toks);
					for (const key of [...gone]) {
						if (app === undefined || key.startsWith(`${app}|`)) {
							gone.delete(key);
						}
					}
					continue;
				}
				if (head === 'deletesetting' && toks[1]?.rawText !== '=') {
					const args = settingArguments(toks);
					const [app, section, key] = args;
					if (app === undefined || args.length > 3 || (args.length > 1 && section === undefined) || (args.length > 2 && key === undefined)) {
						// A setting the rule cannot name: what it deletes is not known.
						if (app === undefined) {
							gone.clear();
						} else {
							for (const fact of [...gone]) {
								if (fact.startsWith(`${app}|`)) {
									gone.delete(fact);
								}
							}
						}
						continue;
					}
					const covered = gone.has(`${app}|*|*`)
						|| (section !== undefined && gone.has(`${app}|${section}|*`))
						|| (key !== undefined && gone.has(`${app}|${section}|${key}`));
					if (covered && !resumeNext) {
						const what = key !== undefined ? 'key' : section !== undefined ? 'section' : 'application';
						push('runtimeArgumentValue', `DeleteSetting finds no such ${what}: this procedure deleted it above. This will raise Run-time error '5': Invalid procedure call or argument.`, { start: node.span.start + toks[0].start, end: node.span.start + toks[toks.length - 1].end });
					}
					gone.add(`${app}|${section ?? '*'}|${key ?? '*'}`);
					continue;
				}
				// Assignments, getters and expressions can invoke code that restores
				// settings too. Keep facts only across modeled setting operations.
				gone.clear();
			}
		};
		run(member.body);
	}
}
