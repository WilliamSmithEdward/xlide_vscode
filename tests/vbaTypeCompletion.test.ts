import { describe, it, expect } from 'vitest';
import {
	resolveTypeCompletions,
	type ProjectTypeName,
	type TypeCompletionContext,
} from '../src/analyzer';

/** Offset just past the end of the first occurrence of `marker`. */
function endOf(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function names(src: string, marker: string, ctx: TypeCompletionContext = {}): string[] {
	return resolveTypeCompletions(src, endOf(src, marker), ctx).map((t) => t.name);
}

describe('type-position detection', () => {
	it('offers types immediately after "As "', () => {
		const result = names('Dim ws As ', 'As ');
		expect(result).toContain('Long');
		expect(result).toContain('String');
		expect(result).toContain('Worksheet');
		expect(result).toContain('Workbook');
	});

	it('offers only creatable project types after "As New "', () => {
		const result = names('Dim c As New ', 'As New ', {
			projectTypes: [
				{ name: 'Person', kind: 'class' },
				{ name: 'CustomerForm', kind: 'userform' },
				{ name: 'Sheet1', kind: 'document' },
				{ name: 'Status', kind: 'enum' },
				{ name: 'TPoint', kind: 'userType' },
			],
		});
		expect(result).toEqual(['Person', 'CustomerForm']);
		expect(result).not.toContain('Long');
		expect(result).not.toContain('Worksheet');
	});

	it('offers creatable project types after expression-level "New"', () => {
		const result = resolveTypeCompletions(
			'Sub T()\n    Set p = New Pe\nEnd Sub\n',
			endOf('Sub T()\n    Set p = New Pe\nEnd Sub\n', 'New Pe'),
			{
				projectTypes: [
					{ name: 'Person', kind: 'class' },
					{ name: 'CustomerForm', kind: 'userform' },
					{ name: 'Sheet1', kind: 'document' },
					{ name: 'Status', kind: 'enum' },
					{ name: 'TPoint', kind: 'userType' },
				],
			},
		);
		expect(result.map((t) => t.name)).toEqual(['Person']);
	});

	it('does not offer non-creatable types after expression-level "New"', () => {
		const result = names('Sub T()\n    Set p = New ', 'New ', {
			projectTypes: [
				{ name: 'Person', kind: 'class' },
				{ name: 'CustomerForm', kind: 'userform' },
				{ name: 'Sheet1', kind: 'document' },
				{ name: 'Status', kind: 'enum' },
				{ name: 'TPoint', kind: 'userType' },
			],
		});
		expect(result).toEqual(['Person', 'CustomerForm']);
		expect(result).not.toContain('Long');
		expect(result).not.toContain('Worksheet');
	});

	it('filters by the partial type text typed', () => {
		const result = names('Dim x As Lo', 'As Lo');
		expect(result).toContain('Long');
		expect(result).toContain('LongLong');
		expect(result).toContain('LongPtr');
		expect(result).not.toContain('String');
		expect(result).not.toContain('Workbook');
	});

	it('is case-insensitive on the partial', () => {
		expect(names('Dim x As wOrK', 'As wOrK')).toContain('Workbook');
	});

	it('offers types for a function return type', () => {
		const result = names('Function F() As ', 'As ');
		expect(result).toContain('Long');
		expect(result).toContain('Range');
	});

	it('offers types for a parameter declaration', () => {
		const result = names('Sub S(ByVal x As ', 'As ');
		expect(result).toContain('Long');
	});

	it('offers types after a comma in a Dim list', () => {
		const result = names('Dim a As Long, b As ', 'b As ');
		expect(result).toContain('String');
	});
});

describe('non-type positions return nothing', () => {
	it('does not offer types before "As" is typed', () => {
		expect(resolveTypeCompletions('Dim ws ', 7)).toEqual([]);
	});

	it('does not offer types at the start of a line', () => {
		expect(resolveTypeCompletions('Dim ', 4)).toEqual([]);
	});

	it('does not treat an identifier named like a keyword as As', () => {
		// "AsName" tokenizes as a single identifier, not the As keyword.
		expect(resolveTypeCompletions('Dim AsName', 10)).toEqual([]);
	});

	it('does not offer types inside a comment', () => {
		expect(resolveTypeCompletions("' Dim x As ", 11)).toEqual([]);
	});
});

describe('project-defined types', () => {
	const projectTypes: ProjectTypeName[] = [
		{ name: 'Customer', kind: 'class' },
		{ name: 'TPoint', kind: 'userType' },
		{ name: 'Color', kind: 'enum' },
	];

	it('includes user classes, types, and enums', () => {
		const result = names('Dim x As ', 'As ', { projectTypes });
		expect(result).toContain('Customer');
		expect(result).toContain('TPoint');
		expect(result).toContain('Color');
	});

	it('lists project types before built-ins', () => {
		const result = names('Dim x As ', 'As ', { projectTypes });
		expect(result.indexOf('Customer')).toBeLessThan(result.indexOf('Long'));
	});

	it('lets a user type shadow a built-in of the same name', () => {
		const result = resolveTypeCompletions('Dim x As ', endOf('Dim x As ', 'As '), {
			projectTypes: [{ name: 'Long', kind: 'class' }],
		});
		const longs = result.filter((t) => t.name.toLowerCase() === 'long');
		expect(longs).toHaveLength(1);
		expect(longs[0].kind).toBe('class');
	});

	it('keeps colliding project type names generic', () => {
		const result = resolveTypeCompletions('Dim x As ', endOf('Dim x As ', 'As '), {
			projectTypes: [
				{ name: 'Status', kind: 'class' },
				{ name: 'Status', kind: 'enum' },
			],
		});
		const status = result.find((t) => t.name === 'Status');
		expect(status?.kind).toBe('ambiguous');
		expect(status?.detail).toBe('Ambiguous project type');
	});

	it('keeps duplicate project type names generic even when their kind matches', () => {
		const result = resolveTypeCompletions('Dim x As ', endOf('Dim x As ', 'As '), {
			projectTypes: [
				{ name: 'Payload', kind: 'userType' },
				{ name: 'Payload', kind: 'userType' },
			],
		});
		const payload = result.find((t) => t.name === 'Payload');
		expect(payload?.kind).toBe('ambiguous');
		expect(payload?.detail).toBe('Ambiguous project type');
	});

	it('tags each candidate with an origin detail', () => {
		const result = resolveTypeCompletions('Dim x As ', endOf('Dim x As ', 'As '), {
			projectTypes,
		});
		const byName = new Map(result.map((t) => [t.name, t]));
		expect(byName.get('Long')?.detail).toBe('VBA type');
		expect(byName.get('Worksheet')?.detail).toBe('Excel type');
		expect(byName.get('Customer')?.detail).toBe('Class');
		expect(byName.get('Color')?.detail).toBe('Enum');
	});

	it('includes rendered documentation for documented project types', () => {
		const result = resolveTypeCompletions('Dim x As ', endOf('Dim x As ', 'As '), {
			projectTypes: [
				{
					name: 'Person',
					kind: 'class',
					doc: {
						summary: 'Represents a person.',
						remarks: 'Stored in the workbook domain model.',
						params: [],
						source: 'inline',
					},
				},
			],
		});
		const person = result.find((t) => t.name === 'Person');
		expect(person?.documentation).toContain('Represents a person.');
		expect(person?.documentation).toContain('Stored in the workbook domain model.');
	});
});

describe('verified candidate set', () => {
	it('does not offer Decimal (not directly declarable in VBA)', () => {
		expect(names('Dim x As ', 'As ')).not.toContain('Decimal');
	});

	it('offers the core VBA primitive types', () => {
		const result = names('Dim x As ', 'As ');
		for (const t of [
			'Boolean',
			'Byte',
			'Currency',
			'Date',
			'Double',
			'Integer',
			'Long',
			'Object',
			'Single',
			'String',
			'Variant',
		]) {
			expect(result).toContain(t);
		}
	});
});
