import { ProjectIndex } from '../symbols/projectIndex';
import { tokenizeCached } from '../lexer/tokenize';
import { firstTokenAtOrAfter, tokenName } from '../lexer/tokenHelpers';
import type { ProcedureNode } from '../parser/nodes';
import type { CallSite } from './callSites';

/** A query-local source index: textual call candidates must bind to the changed procedure. */
export function procedureCallBinding(
	source: string,
	moduleName: string,
	procedure: ProcedureNode,
	others: Readonly<Record<string, string>>,
): (callerName: string, callerSource: string, site: CallSite) => boolean {
	let project: ProjectIndex | undefined;
	const tokensByModule = new Map<string, ReturnType<typeof tokenizeCached>>();
	return (callerName, callerSource, site) => {
		if (!project) {
			project = new ProjectIndex();
			project.setModule({moduleName, moduleKind: 'standard', source});
			for (const [otherName, otherSource] of Object.entries(others)) {
				if (otherName.toLowerCase() !== moduleName.toLowerCase()) {
					project.setModule({moduleName: otherName, moduleKind: 'standard', source: otherSource});
				}
			}
		}
		let tokens = tokensByModule.get(callerName);
		if (!tokens) {
			tokens = tokenizeCached(callerSource).filter(token => token.kind !== 'comment');
			tokensByModule.set(callerName, tokens);
		}
		let index = firstTokenAtOrAfter(tokens, site.offset);
		if (tokens[index - 1]?.kind === 'bracketedIdentifier' && tokens[index - 1].end > site.offset) { index--; }
		const token = tokens[index];
		if (!token || tokenName(token)?.toLowerCase() !== procedure.name.toLowerCase()) { return false; }
		const before = tokens[index - 1];
		const next = tokens[index + 1];
		if (next?.rawText === ':=' || next?.kind === 'colon' && (!before || before.kind === 'newline' || before.kind === 'colon') || before?.rawText === '!'
			|| /^(GoTo|GoSub|Resume|AddressOf)$/i.test(before?.rawText ?? '')) { return false; }
		if (before?.rawText === '.') {
			// Only an unshadowed owning-module qualifier denotes this standard-module
			// procedure. An object member or longer receiver chain is a separate binding.
			const receiver = tokens[index - 2];
			const receiverName = receiver && tokenName(receiver);
			if (receiverName?.toLowerCase() !== moduleName.toLowerCase()
				|| ['.', '!'].includes(tokens[index - 3]?.rawText)
				|| /^(GoTo|GoSub|Resume|AddressOf)$/i.test(tokens[index - 3]?.rawText ?? '')) { return false; }
			const binding = project.resolveBareIdentifier(callerName, receiverName, receiver.start, 'memberReceiver');
			if (binding.scope !== 'unresolved') { return false; }
			const owner = project.getModule(moduleName)!;
			const target = owner.root.children?.find(symbol => symbol.fullSpan.start === procedure.span.start && symbol.name.toLowerCase() === procedure.name.toLowerCase());
			return !!target && (callerName.toLowerCase() === moduleName.toLowerCase() || target.visibility !== 'Private');
		}
		const call = next?.rawText === '(' || before === undefined || before.kind === 'newline' || before.kind === 'colon'
			|| /^(Then|Else|Call)$/i.test(before.rawText);
		const binding = project.resolveBareIdentifier(callerName, procedure.name, token.start, call ? 'call' : 'expression');
		return binding.scope !== 'ambiguous' && binding.definitions.length === 1
			&& binding.definitions[0].moduleName.toLowerCase() === moduleName.toLowerCase()
			&& binding.definitions[0].fullSpan.start === procedure.span.start;
	};
}
