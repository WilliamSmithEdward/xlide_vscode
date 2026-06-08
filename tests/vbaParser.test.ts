// Parser fixtures for the ground-up VBA analyzer (Phase 3).
//
// Verified against MS-VBAL.pdf, v20250520. Each describe block notes the spec
// section that justifies the construct under test. See
// docs/spec/MS-VBAL.verification-map.md for the full Feature->Section->Status map.

import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import {
	ConditionalDirectiveNode,
	EnumNode,
	ForBlockNode,
	ModuleMember,
	ProcedureNode,
	TypeNode,
	VariableGroupNode,
} from '../src/analyzer/parser/nodes';

function procedures(source: string): ProcedureNode[] {
	return parseModule(source).members.filter(
		(m): m is ProcedureNode => m.kind === 'Procedure',
	);
}

function memberKinds(source: string): ModuleMember['kind'][] {
	return parseModule(source).members.map((m) => m.kind);
}

describe('parseModule - module header (MS-VBAL 4.2)', () => {
	it('parses Attribute lines into attribute nodes', () => {
		const m = parseModule('Attribute VB_Name = "Module1"\n');
		expect(m.members).toHaveLength(1);
		const attr = m.members[0];
		expect(attr.kind).toBe('Attribute');
		if (attr.kind === 'Attribute') {
			expect(attr.name).toBe('VB_Name');
			expect(attr.valueRaw).toBe('"Module1"');
		}
	});

	it('keeps dotted member Attribute names intact', () => {
		const source = 'Attribute Value.VB_UserMemId = 0\n';
		const m = parseModule(source);
		const attr = m.members[0];
		expect(attr.kind).toBe('Attribute');
		if (attr.kind === 'Attribute') {
			expect(attr.name).toBe('Value.VB_UserMemId');
			expect(source.slice(attr.nameSpan.start, attr.nameSpan.end)).toBe(
				'Value.VB_UserMemId',
			);
			expect(attr.valueRaw).toBe('0');
		}
	});

	it('attaches exported member Attribute lines in the procedure metadata slot', () => {
		const source =
			'Public Property Get NewEnum() As IUnknown\n' +
			'Attribute NewEnum.VB_UserMemId = -4\n' +
			'    Set NewEnum = Nothing\n' +
			'End Property\n';
		const proc = procedures(source)[0];

		expect(proc.attributes).toHaveLength(1);
		expect(proc.attributes?.[0].name).toBe('NewEnum.VB_UserMemId');
		expect(proc.body).toHaveLength(1);
	});

	it('detects a class module from VB_Exposed attribute', () => {
		const m = parseModule(
			'Attribute VB_Name = "Class1"\nAttribute VB_Exposed = False\n',
		);
		expect(m.moduleKind).toBe('class');
	});
});

describe('parseModule - option directives (MS-VBAL 5.2.1)', () => {
	it('parses Option Explicit', () => {
		const m = parseModule('Option Explicit\n');
		const opt = m.members[0];
		expect(opt.kind).toBe('Option');
		if (opt.kind === 'Option') {
			expect(opt.optionText).toBe('Explicit');
		}
	});

	it('parses Option Compare Text with canonical casing', () => {
		const m = parseModule('option compare text\n');
		const opt = m.members[0];
		if (opt.kind === 'Option') {
			expect(opt.optionText).toBe('Compare Text');
		}
	});
});

describe('parseModule - conditional compilation directives (MS-VBAL 3.4)', () => {
	it('parses #Const compiler constants at module level', () => {
		const m = parseModule('#Const Win64Build = True\n');
		const directive = m.members[0];
		expect(directive.kind).toBe('ConditionalDirective');
		if (directive.kind === 'ConditionalDirective') {
			expect(directive.directiveKind).toBe('Const');
			expect(directive.name).toBe('Win64Build');
			expect(directive.valueRaw).toBe('True');
		}
	});

	it('parses conditional Declare branches in source order', () => {
		const m = parseModule(
			'#If VBA7 Then\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'#Else\n' +
			'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
			'#End If\n',
		);
		expect(m.members.map((member) => member.kind)).toEqual([
			'ConditionalDirective',
			'Declare',
			'ConditionalDirective',
			'Declare',
			'ConditionalDirective',
		]);
		const directives = m.members.filter(
			(member): member is ConditionalDirectiveNode => member.kind === 'ConditionalDirective',
		);
		expect(directives.map((directive) => directive.directiveKind)).toEqual(['If', 'Else', 'EndIf']);
		expect(directives[0].conditionRaw).toBe('VBA7');
	});

	it('parses #ElseIf and #EndIf spelling variants', () => {
		const m = parseModule('#If VBA7 Then\n#ElseIf Win64 Then\n#EndIf\n');
		const directives = m.members.filter(
			(member): member is ConditionalDirectiveNode => member.kind === 'ConditionalDirective',
		);
		expect(directives.map((directive) => directive.directiveKind)).toEqual([
			'If',
			'ElseIf',
			'EndIf',
		]);
		expect(directives[1].conditionRaw).toBe('Win64');
	});

	it('keeps procedure-scope directives in the procedure body', () => {
		const proc = procedures(
			'Sub T()\n' +
			'    #If DEBUGGING Then\n' +
			'    Debug.Print "on"\n' +
			'    #End If\n' +
			'End Sub\n',
		)[0];
		expect(proc.body.map((node) => node.kind)).toEqual([
			'ConditionalDirective',
			'Statement',
			'ConditionalDirective',
		]);
	});
});

describe('parseModule - procedures (MS-VBAL 5.3.1 / 5.3.2)', () => {
	it('extracts a simple Sub with no parameters', () => {
		const procs = procedures('Sub Foo()\nEnd Sub\n');
		expect(procs).toHaveLength(1);
		expect(procs[0].procKind).toBe('Sub');
		expect(procs[0].name).toBe('Foo');
		expect(procs[0].closed).toBe(true);
		expect(procs[0].params).toHaveLength(0);
	});

	it('captures visibility and Static modifiers', () => {
		const procs = procedures('Private Static Sub Foo()\nEnd Sub\n');
		expect(procs[0].modifiers).toEqual(['Private', 'Static']);
	});

	it('parses a Function return type', () => {
		const procs = procedures('Public Function Add(a As Long, b As Long) As Long\nEnd Function\n');
		expect(procs[0].procKind).toBe('Function');
		expect(procs[0].returnType).toBe('Long');
		expect(procs[0].params.map((p) => p.name)).toEqual(['a', 'b']);
		expect(procs[0].params.map((p) => p.asType)).toEqual(['Long', 'Long']);
	});

	it('parses Property Get / Let / Set kinds', () => {
		const src =
			'Property Get Name() As String\nEnd Property\n' +
			'Property Let Name(value As String)\nEnd Property\n' +
			'Property Set Ref(value As Object)\nEnd Property\n';
		const procs = procedures(src);
		expect(procs.map((p) => p.procKind)).toEqual([
			'PropertyGet',
			'PropertyLet',
			'PropertySet',
		]);
	});

	it('parses Optional, ByVal, ByRef and ParamArray parameter markers', () => {
		const procs = procedures(
			'Sub F(ByVal a As Long, ByRef b As String, Optional c As Long = 5, ParamArray rest())\nEnd Sub\n',
		);
		const p = procs[0].params;
		expect(p[0].byVal).toBe(true);
		expect(p[1].byRef).toBe(true);
		expect(p[2].optional).toBe(true);
		expect(p[2].defaultRaw).toBe('5');
		expect(p[3].paramArray).toBe(true);
		expect(p[3].isArray).toBe(true);
	});
});

describe('parseModule - declarations (MS-VBAL 5.2.3 / 5.2.4)', () => {
	it('parses a Dim with multiple declarators', () => {
		const m = parseModule('Dim a As Long, b As String\n');
		const group = m.members[0] as VariableGroupNode;
		expect(group.kind).toBe('VariableGroup');
		expect(group.isConst).toBe(false);
		expect(group.declarations.map((d) => d.name)).toEqual(['a', 'b']);
		expect(group.declarations.map((d) => d.asType)).toEqual(['Long', 'String']);
	});

	it('parses fixed-length String declarations as String plus a fixed length', () => {
		const m = parseModule(
			'Dim fixedName As String * 20\n' +
				'Private Type Header\n' +
				'    Code As String * 4\n' +
				'End Type\n',
		);
		const group = m.members[0] as VariableGroupNode;
		expect(group.declarations[0].asType).toBe('String');
		expect(group.declarations[0].fixedLength).toBe('20');

		const type = m.members[1] as TypeNode;
		expect(type.fields[0].asType).toBe('String');
		expect(type.fields[0].fixedLength).toBe('4');
	});

	it('parses legacy type-declaration suffixes as base names and inferred types', () => {
		const m = parseModule(
			'Public Function GetName$()\nEnd Function\n' +
				'Public Sub Demo(ByVal label$)\n' +
				'    Dim total&, name$, price@, ratio#, flag%\n' +
				'End Sub\n' +
				'Private Type Header\n' +
				'    Code$\n' +
				'End Type\n',
		);
		const fn = m.members[0] as ProcedureNode;
		expect(fn.name).toBe('GetName');
		expect(fn.typeSuffix).toBe('$');
		expect(fn.returnType).toBe('String');

		const sub = m.members[1] as ProcedureNode;
		expect(sub.params[0].name).toBe('label');
		expect(sub.params[0].asType).toBe('String');

		const group = sub.body[0] as VariableGroupNode;
		expect(group.declarations.map((d) => d.name)).toEqual([
			'total',
			'name',
			'price',
			'ratio',
			'flag',
		]);
		expect(group.declarations.map((d) => d.asType)).toEqual([
			'Long',
			'String',
			'Currency',
			'Double',
			'Integer',
		]);

		const type = m.members[2] as TypeNode;
		expect(type.fields[0].name).toBe('Code');
		expect(type.fields[0].asType).toBe('String');
	});

	it('parses a module variable declared with only a visibility modifier', () => {
		const m = parseModule('Public Counter As Long\n');
		const group = m.members[0] as VariableGroupNode;
		expect(group.modifier).toBe('Public');
		expect(group.declarations[0].name).toBe('Counter');
	});

	it('distinguishes fixed and dynamic array declarators', () => {
		const m = parseModule('Dim fixed(1 To 3) As Long, dynamic() As Long\n');
		const group = m.members[0] as VariableGroupNode;
		expect(group.declarations[0].isArray).toBe(true);
		expect(group.declarations[0].arrayBounds).toBe('1 To 3');
		expect(group.declarations[1].isArray).toBe(true);
		expect(group.declarations[1].arrayBounds).toBeUndefined();
	});

	it('parses Const declarations', () => {
		const m = parseModule('Const Pi As Double = 3.14159\n');
		const group = m.members[0] as VariableGroupNode;
		expect(group.isConst).toBe(true);
		expect(group.declarations[0].name).toBe('Pi');
		expect(group.declarations[0].asType).toBe('Double');
		expect(group.declarations[0].defaultRaw).toBe('3.14159');
	});

	it('flags WithEvents and New declarations', () => {
		const m = parseModule('Private WithEvents src As Object\nDim x As New Collection\n');
		const g0 = m.members[0] as VariableGroupNode;
		const g1 = m.members[1] as VariableGroupNode;
		expect(g0.withEvents).toBe(true);
		expect(g1.declarations[0].isNew).toBe(true);
	});

	it('parses external Declare metadata', () => {
		const m = parseModule(
			'Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal ClassName As String, ByVal WindowName As String) As LongPtr\n',
		);
		const decl = m.members[0];
		expect(decl.kind).toBe('Declare');
		if (decl.kind === 'Declare') {
			expect(decl.name).toBe('FindWindow');
			expect(decl.isFunction).toBe(true);
			expect(decl.visibility).toBe('Public');
			expect(decl.ptrSafe).toBe(true);
			expect(decl.libName).toBe('user32');
			expect(decl.aliasName).toBe('FindWindowA');
			expect(decl.returnType).toBe('LongPtr');
			expect(decl.params.map((p) => `${p.name}:${p.asType}:${p.byVal}`)).toEqual([
				'ClassName:String:true',
				'WindowName:String:true',
			]);
		}
	});

	it('parses continued external Declare parameter lists', () => {
		const m = parseModule(
			'Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _\n' +
			'    Destination As Any, _\n' +
			'    Source As Any, _\n' +
			'    ByVal Length As LongPtr)\n',
		);
		const decl = m.members[0];
		expect(decl.kind).toBe('Declare');
		if (decl.kind === 'Declare') {
			expect(decl.name).toBe('CopyMemory');
			expect(decl.aliasName).toBe('RtlMoveMemory');
			expect(decl.params.map((p) => `${p.name}:${p.asType}:${p.byVal}`)).toEqual([
				'Destination:Any:false',
				'Source:Any:false',
				'Length:LongPtr:true',
			]);
		}
	});
});

describe('parseModule - Type and Enum (MS-VBAL 5.2.3.3 / 5.2.3.4)', () => {
	it('parses a Type block with fields', () => {
		const m = parseModule('Type TPoint\n    X As Long\n    Y As Long\nEnd Type\n');
		const t = m.members[0] as TypeNode;
		expect(t.kind).toBe('Type');
		expect(t.name).toBe('TPoint');
		expect(t.closed).toBe(true);
		expect(t.fields.map((f) => f.name)).toEqual(['X', 'Y']);
	});

	it('parses an Enum block with members', () => {
		const m = parseModule('Public Enum Color\n    Red\n    Green = 2\n    Blue\nEnd Enum\n');
		const e = m.members[0] as EnumNode;
		expect(e.kind).toBe('Enum');
		expect(e.name).toBe('Color');
		expect(e.visibility).toBe('Public');
		expect(e.members.map((x) => x.name)).toEqual(['Red', 'Green', 'Blue']);
		expect(e.members.map((x) => x.valueRaw)).toEqual([undefined, '2', undefined]);
	});
});

describe('parseModule - block statements (MS-VBAL 5.4)', () => {
	it('nests If/For/With/Do/Select/While blocks', () => {
		const src = [
			'Sub F()',
			'    If x Then',
			'        For i = 1 To 10',
			'            With obj',
			'                Do',
			'                    Select Case i',
			'                        Case 1',
			'                    End Select',
			'                Loop',
			'            End With',
			'        Next',
			'    End If',
			'    While y',
			'    Wend',
			'End Sub',
		].join('\n');
		const procs = procedures(src);
		expect(procs[0].closed).toBe(true);
		expect(parseModule(src).diagnostics).toHaveLength(0);
	});

	it('treats a single-line If as a statement, not a block', () => {
		const src = 'Sub F()\n    If x Then y = 1\nEnd Sub\n';
		const m = parseModule(src);
		expect(m.diagnostics).toHaveLength(0);
		const proc = m.members[0] as ProcedureNode;
		expect(proc.body.every((n) => n.kind !== 'IfBlock')).toBe(true);
	});

	it('recognizes block statements after numeric line labels', () => {
		const src =
			'Sub F()\n' +
			'10 If ready Then\n' +
			'20 End If\n' +
			'30 For i = 1 To 3\n' +
			'40 Next i\n' +
			'50 For Each item In items\n' +
			'60 Next item\n' +
			'70 Do\n' +
			'80 Loop\n' +
			'90 While flag\n' +
			'100 Wend\n' +
			'110 With obj\n' +
			'120 .Name = "x"\n' +
			'130 End With\n' +
			'140 Select Case x\n' +
			'150 Case 1\n' +
			'160 End Select\n' +
			'End Sub\n';
		const m = parseModule(src);
		const proc = m.members[0] as ProcedureNode;
		expect(m.diagnostics).toHaveLength(0);
		expect(proc.body.map((node) => node.kind)).toEqual([
			'IfBlock',
			'ForBlock',
			'ForBlock',
			'DoBlock',
			'WhileBlock',
			'WithBlock',
			'SelectBlock',
		]);
	});

	it('distinguishes For Each from For', () => {
		const src = 'Sub F()\n    For Each item In coll\n    Next\nEnd Sub\n';
		const proc = parseModule(src).members[0] as ProcedureNode;
		const block = proc.body.find((n) => n.kind === 'ForBlock');
		expect(block && block.kind === 'ForBlock' && block.each).toBe(true);
	});

	it('captures simple For and Next control variables', () => {
		const src =
			'Sub F()\n' +
			'    For i = 1 To 3\n' +
			'    Next j\n' +
			'    For Each item In coll\n' +
			'    Next item\n' +
			'End Sub\n';
		const proc = parseModule(src).members[0] as ProcedureNode;
		const blocks = proc.body.filter((n) => n.kind === 'ForBlock') as ForBlockNode[];
		expect(blocks.map((block) => block.controlVariable)).toEqual(['i', 'item']);
		expect(blocks.map((block) => block.nextVariable)).toEqual(['j', 'item']);
		expect(blocks[1].sourceExpression).toBe('coll');
		expect(
			blocks[1].sourceExpressionSpan
				? src.slice(blocks[1].sourceExpressionSpan.start, blocks[1].sourceExpressionSpan.end)
				: undefined,
		).toBe('coll');
	});
});

describe('parseModule - error recovery (Phase 3 acceptance)', () => {
	it('never throws on malformed input', () => {
		expect(() => parseModule('Sub (((\nEnd If\nProperty\n#If\n')).not.toThrow();
	});

	it('reports a missing End Sub but still extracts the procedure', () => {
		const m = parseModule('Sub Broken()\n    x = 1\n');
		const proc = m.members[0] as ProcedureNode;
		expect(proc.name).toBe('Broken');
		expect(proc.closed).toBe(false);
		expect(m.diagnostics.some((d) => /missing End Sub/i.test(d.message))).toBe(true);
	});

	it('recovers the next procedure when an End is forgotten', () => {
		const m = parseModule('Sub A()\n    x = 1\nSub B()\nEnd Sub\n');
		const procs = procedures('Sub A()\n    x = 1\nSub B()\nEnd Sub\n');
		expect(procs.map((p) => p.name)).toEqual(['A', 'B']);
		expect(m.members.filter((x) => x.kind === 'Procedure')).toHaveLength(2);
	});

	it('reports a stray block terminator', () => {
		const m = parseModule('Sub F()\n    End If\nEnd Sub\n');
		expect(m.diagnostics.some((d) => /Unexpected 'End If'/.test(d.message))).toBe(true);
	});

	it('reports an unclosed block inside a procedure', () => {
		const m = parseModule('Sub F()\n    If x Then\nEnd Sub\n');
		expect(m.diagnostics.some((d) => /missing End If/i.test(d.message))).toBe(true);
	});

	it('keeps a procedure closed after invalid nested Type and Enum blocks', () => {
		const src =
			'Sub F()\n' +
			'    Type TInside\n' +
			'        Value As Long\n' +
			'    End Type\n' +
			'    Enum EInside\n' +
			'        A = 1\n' +
			'    End Enum\n' +
			'End Sub\n';
		const m = parseModule(src);
		const proc = m.members[0] as ProcedureNode;

		expect(m.members).toHaveLength(1);
		expect(proc.kind).toBe('Procedure');
		expect(proc.closed).toBe(true);
		expect(proc.body.map((node) => node.kind)).toEqual(['Statement', 'Statement']);
		expect(src.slice(proc.span.start, proc.span.end)).toContain('End Sub');
	});

	it('reports an unclosed Type block', () => {
		const m = parseModule('Type T\n    X As Long\n');
		expect(m.diagnostics.some((d) => /missing End Type/i.test(d.message))).toBe(true);
	});
});

describe('parseModule - spans', () => {
	it('gives every node an absolute source span', () => {
		const src = 'Sub Foo()\nEnd Sub\n';
		const proc = parseModule(src).members[0] as ProcedureNode;
		expect(proc.span.start).toBe(0);
		expect(src.slice(proc.span.start, proc.span.end)).toContain('Sub Foo()');
		expect(src.slice(proc.span.start, proc.span.end)).toContain('End Sub');
	});

	it('orders members as written', () => {
		const kinds = memberKinds(
			'Attribute VB_Name = "M"\nOption Explicit\nDim g As Long\nSub Foo()\nEnd Sub\n',
		);
		expect(kinds).toEqual(['Attribute', 'Option', 'VariableGroup', 'Procedure']);
	});
});
