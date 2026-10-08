import { describe, expect, it } from 'vitest';
import { collectTypeNameReferences } from '../src/analyzer/semantic/typeSemanticTokens';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';

// The VBE accepts Type members named As and Implements (oracle case
// reserved_member_name_every_reserved_word_type_fields_compile). Their type
// is the name after the member's own As, and neither line is an Implements
// statement; interfaces whose names start with As still are.
const SOURCE = [
	'Implements As_Foo',
	'Implements Asx',
	'Private Type Rec',
	'    As As Long',
	'    Implements As Long',
	'End Type',
	'',
].join('\n');

describe('Type members named As and Implements', () => {
	it('reference the type after their own As, and implement nothing', () => {
		const references = collectTypeNameReferences(SOURCE)
			.map((ref) => [ref.kind, SOURCE.slice(ref.span.start, ref.span.end), SOURCE.lastIndexOf('\n', ref.span.start) + 1]);
		const lineStart = (line: string): number => SOURCE.indexOf(line);
		expect(references).toEqual([
			['implements', 'As_Foo', 0],
			['implements', 'Asx', lineStart('Implements Asx')],
			['declaration', 'Long', lineStart('    As As Long')],
			['declaration', 'Long', lineStart('    Implements As Long')],
		]);
	});

	it('leave the class implementing only its real interfaces', () => {
		const index = new ProjectIndex();
		index.setModule({ moduleName: 'Widget', moduleKind: 'class', source: SOURCE });
		expect(index.moduleImplementsList('Widget')).toEqual(['As_Foo', 'Asx']);
	});
});
