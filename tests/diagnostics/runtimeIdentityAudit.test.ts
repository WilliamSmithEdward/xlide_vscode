import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

const errors = (source: string, options: Parameters<typeof analyzeModule>[1] = {}) => analyzeModule(source, options).filter(d => d.severity === 'error');
const wrap = (body: string, extra = '') => `Option Explicit\nSub Main()\n${body}\nEnd Sub\n${extra}`;

describe('runtime identity audit', () => {
	it('does not apply Word name rules to a custom late-bound document', () => {
		expect(errors(wrap('Dim d As Object\nSet d = CreateObject("Acme.Document")\nd.Variables.Add "x", 1\nd.Variables.Add "x", 2'), { hostModel: getWordObjectModel() })).toEqual([]);
	});
	it('does not use JavaScript case folding for locale-sensitive dictionary keys', () => {
		expect(errors(wrap('Dim d As Object\nSet d = CreateObject("Scripting.Dictionary")\nd.CompareMode = 1\nd.Add "I", 1\nd.Add "i", 2'))).toEqual([]);
	});
	it.each(['Collection', 'Object'])('does not infer contents after a handled failed Add: %s', type => {
		const create = type === 'Collection' ? 'New Collection' : 'CreateObject("Scripting.Dictionary")';
		const add = type === 'Collection' ? 'c.Add 1, "k"' : 'c.Add "k", 1';
		expect(errors(wrap(`Dim c As ${type}\nSet c = ${create}\n${add}\nOn Error Resume Next\n${add}\nOn Error GoTo 0\nDebug.Print c.Count`))).toEqual([]);
	});
	it('forgets workbook protection before a bare getter runs', () => {
		const source = 'Option Explicit\nDim wb As Workbook\nSub Main()\nDim ws As Worksheet\nSet wb = ThisWorkbook\nSet ws = wb.Worksheets(1)\nwb.Protect "pw", True\nws.Name = NewName\nEnd Sub\nProperty Get NewName() As String\nwb.Unprotect "pw"\nNewName = "Renamed"\nEnd Property';
		expect(errors(source)).toEqual([]);
	});
	it('forgets a file mode before a block condition reopens the handle', () => {
		expect(errors(wrap('Open "xlide-audit.tmp" For Input As #1\nIf Reopen() Then\nPrint #1, "x"\nEnd If', 'Function Reopen() As Boolean\nClose #1\nOpen "xlide-audit.tmp" For Output As #1\nReopen = True\nEnd Function'))).toEqual([]);
	});
	it('does not infer a runtime class for a shared module field', () => {
		const source = `Option Explicit\nDim o As Object\nSub Main()\nSet o = New Collection\nReplaceObject\no.RemoveAll\nEnd Sub\nSub ReplaceObject()\nSet o = CreateObject("Scripting.Dictionary")\nEnd Sub`;
		expect(errors(source)).toEqual([]);
	});
	it('does not infer a class from an assignment that may fail', () => {
		expect(errors(wrap('Dim o As Object\nSet o = CreateObject("Scripting.Dictionary")\nOn Error Resume Next\nSet o = New Collection\nOn Error GoTo 0\no.RemoveAll'))).toEqual([]);
	});
	it('forgets a RegExp pattern changed through an escaped alias', () => {
		expect(errors(wrap('Dim re As Object\nSet re = CreateObject("VBScript.RegExp")\nSave re\nre.Pattern = "("\nFixPattern\nDebug.Print re.Test("x")', 'Sub Save(re As Object)\nSet saved = re\nEnd Sub\nSub FixPattern()\nsaved.Pattern = "x"\nEnd Sub').replace('Option Explicit', 'Option Explicit\nDim saved As Object'))).toEqual([]);
	});
	it.each(['d.CompareMode = Reset(other)', 'd.Key("a") = Reset(other)'])('forgets other dictionaries during argument evaluation: %s', mutation => {
		const body = `Dim d As Object, other As Object\nSet d = CreateObject("Scripting.Dictionary")\nSet other = CreateObject("Scripting.Dictionary")\nd.Add "a", 1\nother.Add "k", 1\n${mutation}\nother.Add "k", 2`;
		expect(errors(wrap(body, 'Function Reset(d As Object) As Long\nd.RemoveAll\nReset = 0\nEnd Function'))).toEqual([]);
	});
	it('forgets a deleted object passed as the first bare ByRef argument', () => {
		expect(errors(wrap('Dim n As Name\nSet n = ThisWorkbook.Names.Add("audit", "=1")\nn.Delete\nRestore n\nDebug.Print n.Name', 'Sub Restore(ByRef n As Name)\nSet n = ThisWorkbook.Names.Add("audit", "=1")\nEnd Sub'))).toEqual([]);
	});
	it('does not compare names across presentations', () => {
		const body = 'Dim p As Presentation, q As Presentation, a As Slide, b As Slide\nSet p = Presentations.Add\nSet q = Presentations.Add\nSet a = p.Slides.Add(1, ppLayoutBlank)\nSet b = q.Slides.Add(1, ppLayoutBlank)\na.Name = "same"\nb.Name = "same"';
		expect(errors(wrap(body), { hostModel: getPowerPointObjectModel() })).toEqual([]);
	});
	it('retains duplicate literal names on fixed slides in one presentation', () => {
		const body = 'Dim p As Presentation, a As Slide, b As Slide\nSet p = Presentations.Add\nSet a = p.Slides.Add(1, ppLayoutBlank)\nSet b = p.Slides.Add(2, ppLayoutBlank)\na.Name = "same"\nb.Name = "same"';
		expect(errors(wrap(body), { hostModel: getPowerPointObjectModel() }).map(d => d.code)).toEqual(['host-argument-out-of-range']);
	});
});
