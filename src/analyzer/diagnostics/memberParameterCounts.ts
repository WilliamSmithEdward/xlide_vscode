import { runtimeSignatureParameterText, splitSignatureTopLevel } from './typeInference';

/** Counts source accessor/function parameters without treating quoted defaults as separators. */
export function memberParameterCounts(signature: string | undefined): { total: number; required: number } {
	const inner = signature === undefined ? undefined : runtimeSignatureParameterText(signature)?.trim();
	if (!inner) { return { total: 0, required: 0 }; }
	const parts = splitSignatureTopLevel(inner).map(part => part.trim());
	return { total: parts.length, required: parts.filter(part => !/^(Optional|ParamArray)\b/i.test(part) && !part.startsWith('[')).length };
}
