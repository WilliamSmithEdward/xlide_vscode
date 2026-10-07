import { tokenizeCached } from './analyzer/lexer/tokenize';
import { VBA_RUNTIME_CONSTANTS } from './analyzer/runtime/vbaRuntime';
import { assignmentTargetFromTokens } from './analyzer/completion/assignmentTarget';
import { isColorAssignmentTarget } from './analyzer/completion/assignmentValueCompletion';
import { buildModuleSymbols } from './analyzer/symbols/buildModuleSymbols';
import { parseVbaIntegerLiteral } from './analyzer/constants/integerConstantExpression';
import { isProcedureKind } from './analyzer/symbols/symbolModel';

export interface VbaColorSpan { start: number; end: number; red: number; green: number; blue: number }

const namedColors = new Map(VBA_RUNTIME_CONSTANTS
	.filter(c => c.module === 'ColorConstants' && typeof c.value === 'number')
	.map(c => [c.name.toLowerCase(), Number(c.value)]));

/** VBA stores RGB in a Long with red in the least significant byte. */
function packedColor(start: number, end: number, value: number): VbaColorSpan {
	return { start, end, red: value & 255, green: (value >>> 8) & 255, blue: (value >>> 16) & 255 };
}

function integer(raw: string): number | undefined {
	const value = parseVbaIntegerLiteral(raw);
	return value !== undefined && (!raw.endsWith('%') || (value >= -32768 && value <= 32767)) ? value : undefined;
}

interface ColorBindings {
	module: Set<string>;
	scopes: { start: number; end: number; names: Set<string> }[];
}
const bindingCache: { source: string; bindings: ColorBindings }[] = [];

/** A bounded snapshot index; candidate lookup never walks every declaration. */
export function vbaColorNameShadowedAt(source: string, name: string, offset: number): boolean {
	let bindings = bindingCache.find(entry => entry.source === source)?.bindings;
	if (!bindings) {
		bindings = { module: new Set(), scopes: [] };
		for (const symbol of buildModuleSymbols('Module', 'standard', source).root.children ?? []) {
			bindings.module.add(symbol.name.toLowerCase());
			if (isProcedureKind(symbol.kind)) {
				bindings.scopes.push({ start: symbol.fullSpan.start, end: symbol.fullSpan.end,
					names: new Set((symbol.children ?? []).map(child => child.name.toLowerCase())) });
			} else if (symbol.kind === 'enum') {
				for (const child of symbol.children ?? []) { bindings.module.add(child.name.toLowerCase()); }
			}
		}
		bindings.scopes.sort((a, b) => a.start - b.start);
		bindingCache.unshift({ source, bindings });
		if (bindingCache.length > 2) { bindingCache.pop(); }
	}
	const lower = name.toLowerCase();
	if (bindings.module.has(lower)) { return true; }
	let lo = 0, hi = bindings.scopes.length;
	while (lo < hi) {
		const mid = (lo + hi) >>> 1;
		if (bindings.scopes[mid].start <= offset) { lo = mid + 1; } else { hi = mid; }
	}
	const scope = bindings.scopes[lo - 1];
	return !!scope && offset <= scope.end && scope.names.has(lower);
}

/** Only literal colors are editable; expressions and system OLE colors stay intact. */
export function collectVbaColors(source: string, isCancelled: () => boolean = () => false): VbaColorSpan[] {
	if (isCancelled()) { return []; }
	const tokens = tokenizeCached(source);
	const colors: VbaColorSpan[] = [];
	let statementStart = 0;
	for (let i = 0; i < tokens.length; i++) {
		if ((i & 255) === 0 && isCancelled()) { return []; }
		const token = tokens[i];
		if (token.kind === 'newline' || token.kind === 'colon'
			|| (token.kind === 'keyword' && /^(Then|Else)$/i.test(token.rawText))) {
			statementStart = i + 1;
			continue;
		}
		if (token.kind !== 'identifier' && token.kind !== 'integerLiteral') { continue; }
		const lower = token.rawText.toLowerCase();
		const previous = tokens[i - 1];
		if (token.kind === 'integerLiteral') {
			if (previous?.rawText !== '=') { continue; }
			const next = tokens[i + 1];
			if (next && next.kind !== 'newline' && next.kind !== 'colon' && next.kind !== 'comment'
				&& !(next.kind === 'keyword' && next.rawText.toLowerCase() === 'else')) { continue; }
			const target = assignmentTargetFromTokens(tokens.slice(statementStart, i));
			const value = integer(token.rawText);
			if (target && isColorAssignmentTarget(target) && value !== undefined && value >= 0 && value <= 0xffffff) {
				colors.push(packedColor(token.start, token.end, value));
			}
			continue;
		}
		const named = namedColors.get(lower);
		if (named === undefined && lower !== 'rgb') { continue; }
		let qualifierStart = i;
		while (qualifierStart >= 2 && i - qualifierStart < 4 && tokens[qualifierStart - 1].rawText === '.'
			&& tokens[qualifierStart - 2].kind === 'identifier') { qualifierStart -= 2; }
		if (tokens[qualifierStart - 1]?.rawText === '.') { continue; }
		const qualifier = tokens.slice(qualifierStart, i).map(t => t.rawText).join('');
		const qualifiedVba = /^(VBA\.|ColorConstants\.|VBA\.ColorConstants\.)$/i.test(qualifier);
		if (previous?.rawText === '.' && !qualifiedVba) { continue; }
		if (vbaColorNameShadowedAt(source, qualifiedVba ? tokens[qualifierStart].rawText : lower, token.start)) { continue; }
		if (named !== undefined) {
			colors.push(packedColor(tokens[qualifierStart].start, token.end, named));
			continue;
		}
		if (lower === 'rgb' && (!qualifiedVba || qualifier.toLowerCase() === 'vba.')) {
			const call = tokens.slice(i + 1, i + 8);
			if (call.length !== 7 || call[0].rawText !== '(' || call[2].rawText !== ','
				|| call[4].rawText !== ',' || call[6].rawText !== ')') { continue; }
			const channels = [call[1], call[3], call[5]].map(t => t.kind === 'integerLiteral' ? integer(t.rawText) : undefined);
			if (channels.some(v => v === undefined || v < 0 || v > 255)) { continue; }
			colors.push({ start: tokens[qualifierStart].start, end: call[6].end,
				red: channels[0]!, green: channels[1]!, blue: channels[2]! });
			i += 7;
			continue;
		}
	}
	return colors;
}

export function vbaColorPresentations(red: number, green: number, blue: number): string[] {
	const [r, g, b] = [red, green, blue].map(c => Math.round(Math.max(0, Math.min(1, c)) * 255));
	const packed = r + (g << 8) + (b << 16);
	return [`RGB(${r}, ${g}, ${b})`, `&H${packed.toString(16).toUpperCase().padStart(6, '0')}&`, String(packed)];
}
