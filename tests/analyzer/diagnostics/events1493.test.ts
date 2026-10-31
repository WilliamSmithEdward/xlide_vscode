// Issue #1493, measured through the VBE oracle in Excel, Word, PowerPoint and
// Access 16.0 (build 20527). An untyped event parameter takes the DefType of
// the module that declares the event, and an untyped handler parameter the
// DefType of the handler's module, with a type character winning over both;
// library and UserForm events included. A project Collection that declares
// an Event can be listened to `WithEvents ... As Collection`, and
// `As VBA.Collection` never can. In 64-bit Office a LongPtr parameter takes
// a LongLong handler and the reverse, while an Enum parameter takes Long but
// not LongPtr. The cases are the issue1493_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json; each handler holds a statement,
// since the VBE checks an empty one only when its module runs (issue #195).

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase { id: string; expected: 'accepted' | 'rejected'; host?: string; modules: OracleFixtureModule[] }

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1493_'));

function errors(oracleCase: Pick<OracleCase, 'host' | 'modules'>): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		designerClass: mod.designerClass, implicitMembers: mod.implicitMembers,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${mod.moduleName}: ${diagnostic.code}: ${diagnostic.message}`));
}

describe('events and their handlers (issue #1493)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(164);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			expect(found.length).toBeGreaterThan(0);
		}
	});

	const pulse = (sourceDef: string, listenerDef: string, handler: string): OracleCase['modules'] => [
		{ name: 'Source', type: 'class', source: `Option Explicit\r\n${sourceDef}\r\nPublic Event Pulse(ByVal x)\r\n` },
		{ name: 'Listener', type: 'class', source: `Option Explicit\r\n${listenerDef}\r\nPrivate WithEvents src As Source\r\nPrivate Sub src_Pulse(${handler})\r\n    Debug.Print 1\r\nEnd Sub\r\n` },
	];

	it('names the DefType that typed the handler\'s parameter', () => {
		expect(errors({ modules: pulse('DefLng X', 'DefStr X', 'ByVal x') })).toEqual([
			expect.stringContaining("parameter 1, 'x' is String (by the module's DefType), and the event's is Long"),
		]);
	});

	it('keeps the event as written for hover and completion', () => {
		const project = buildVbaProjectIndex(pulse('DefLng X', '', 'ByVal x As Long').map(oracleModuleInput), undefined, { conditionalCompilation: { projectConstants: {} } });
		const event = project.projectMemberSurfaces('Listener').find((type) => type.name === 'Source')?.members.find((member) => member.name === 'Pulse');
		expect(event?.signature).toBe('Pulse(ByVal x)');
		expect(event?.eventParams?.map((param) => param.type)).toEqual(['Long']);
	});

	it('says why As VBA.Collection and a standard module named Collection are refused', () => {
		const listener = (asType: string): OracleCase['modules'][number] => ({ name: 'Listener', type: 'class', source: `Option Explicit\r\nPrivate WithEvents c As ${asType}\r\n` });
		expect(errors({ modules: [listener('VBA.Collection')] })).toEqual([expect.stringContaining("VBA's Collection, which does not source automation events")]);
		expect(errors({ modules: [listener('[VBA.Collection]')] })).toEqual([expect.stringContaining('User-defined type not defined')]);
		expect(errors({ modules: [{ name: 'Collection', type: 'standard', source: 'Option Explicit\r\n' }, listener('Collection')] }))
			.toEqual([expect.stringContaining('A module is not a valid type')]);
	});
});
