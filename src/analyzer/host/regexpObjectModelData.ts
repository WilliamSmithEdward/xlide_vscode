// Generated from reference/regexp/json by generate-host-object-model.mjs.
// Do not hand-edit: regenerate instead.
//
// Types, aliases and enum constants of VBScript_RegExp_55, read from its registered
// COM type library via pinned pyVBAReference. Types remain
// NON-exhaustive: these snapshots offer and describe visible members.

import type { HostConstant, HostEnum, HostType } from './excelObjectModel';

// The literals live inside a function body so V8 defers parsing and
// evaluating them until the host model is first requested: the extension
// bundle and the analysis worker both load this module at startup, and
// an eagerly evaluated VBScript_RegExp_55 model would cost startup time and heap
// in every session that never opens one of its files.
export interface VBScript_RegExp_55ReferenceData {
	readonly types: Readonly<Record<string, HostType>>;
	readonly aliases: Readonly<Record<string, string>>;
	readonly constants: Readonly<Record<string, HostConstant>>;
	readonly enums: Readonly<Record<string, HostEnum>>;
}

let CACHE: VBScript_RegExp_55ReferenceData | undefined;

export function regexpReferenceData(): VBScript_RegExp_55ReferenceData {
	CACHE ??= {
		types: {
			"VBScript_RegExp_55.IMatch": {"displayName":"IMatch","members":[{"name":"FirstIndex","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Length","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Value","kind":"property","declaredType":"String","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.IMatch2": {"displayName":"IMatch2","members":[{"name":"FirstIndex","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Length","kind":"property","declaredType":"Long","access":"read-only"},{"name":"SubMatches","kind":"property","returns":"VBScript_RegExp_55.SubMatches","declaredType":"Object","access":"read-only"},{"name":"Value","kind":"property","declaredType":"String","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.IMatchCollection": {"displayName":"IMatchCollection","members":[{"name":"Count","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Item","kind":"property","declaredType":"Object","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.IMatchCollection2": {"displayName":"IMatchCollection2","members":[{"name":"Count","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Item","kind":"property","declaredType":"Object","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.IRegExp": {"displayName":"IRegExp","members":[{"name":"Execute","kind":"method","signature":"Execute(sourceString As String) As Object","doc":{"params":[{"name":"sourceString","text":"","type":"String"}],"source":"external"}},{"name":"Global","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"IgnoreCase","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"Pattern","kind":"property","declaredType":"String","access":"read/write"},{"name":"Replace","kind":"method","signature":"Replace(sourceString As String, replaceString As String) As String","doc":{"params":[{"name":"sourceString","text":"","type":"String"},{"name":"replaceString","text":"","type":"String"}],"source":"external"}},{"name":"Test","kind":"method","signature":"Test(sourceString As String) As Boolean","doc":{"params":[{"name":"sourceString","text":"","type":"String"}],"source":"external"}}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.IRegExp2": {"displayName":"IRegExp2","members":[{"name":"Execute","kind":"method","signature":"Execute(sourceString As String) As Object","doc":{"params":[{"name":"sourceString","text":"","type":"String"}],"source":"external"}},{"name":"Global","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"IgnoreCase","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"Multiline","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"Pattern","kind":"property","declaredType":"String","access":"read/write"},{"name":"Replace","kind":"method","signature":"Replace(sourceString As String, replaceVar As Variant) As String","doc":{"params":[{"name":"sourceString","text":"","type":"String"},{"name":"replaceVar","text":"","type":"Variant"}],"source":"external"}},{"name":"Test","kind":"method","signature":"Test(sourceString As String) As Boolean","doc":{"params":[{"name":"sourceString","text":"","type":"String"}],"source":"external"}}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.ISubMatches": {"displayName":"ISubMatches","members":[{"name":"Count","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Item","kind":"property","declaredType":"Variant","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.Match": {"displayName":"Match","members":[{"name":"FirstIndex","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Length","kind":"property","declaredType":"Long","access":"read-only"},{"name":"SubMatches","kind":"property","returns":"VBScript_RegExp_55.SubMatches","declaredType":"Object","access":"read-only"},{"name":"Value","kind":"property","declaredType":"String","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.MatchCollection": {"displayName":"MatchCollection","members":[{"name":"Count","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Item","kind":"property","declaredType":"Object","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.RegExp": {"displayName":"RegExp","members":[{"name":"Execute","kind":"method","signature":"Execute(sourceString As String) As Object","doc":{"params":[{"name":"sourceString","text":"","type":"String"}],"source":"external"}},{"name":"Global","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"IgnoreCase","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"Multiline","kind":"property","declaredType":"Boolean","access":"read/write"},{"name":"Pattern","kind":"property","declaredType":"String","access":"read/write"},{"name":"Replace","kind":"method","signature":"Replace(sourceString As String, replaceVar As Variant) As String","doc":{"params":[{"name":"sourceString","text":"","type":"String"},{"name":"replaceVar","text":"","type":"Variant"}],"source":"external"}},{"name":"Test","kind":"method","signature":"Test(sourceString As String) As Boolean","doc":{"params":[{"name":"sourceString","text":"","type":"String"}],"source":"external"}}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
			"VBScript_RegExp_55.SubMatches": {"displayName":"SubMatches","members":[{"name":"Count","kind":"property","declaredType":"Long","access":"read-only"},{"name":"Item","kind":"property","declaredType":"Variant","access":"read-only"}],"provenance":"VBScript_RegExp_55 5.5 via pyVBAReference a6c39d0617085297e88bec7de0ee8afb7aa0746b"},
		},
		aliases: {"imatch":"VBScript_RegExp_55.IMatch","imatch2":"VBScript_RegExp_55.Match","imatchcollection":"VBScript_RegExp_55.IMatchCollection","imatchcollection2":"VBScript_RegExp_55.MatchCollection","iregexp":"VBScript_RegExp_55.IRegExp","iregexp2":"VBScript_RegExp_55.RegExp","isubmatches":"VBScript_RegExp_55.SubMatches","match":"VBScript_RegExp_55.Match","matchcollection":"VBScript_RegExp_55.MatchCollection","regexp":"VBScript_RegExp_55.RegExp","submatches":"VBScript_RegExp_55.SubMatches","vbscript_regexp_55.imatch2":"VBScript_RegExp_55.Match","vbscript_regexp_55.imatchcollection2":"VBScript_RegExp_55.MatchCollection","vbscript_regexp_55.iregexp2":"VBScript_RegExp_55.RegExp","vbscript_regexp_55.isubmatches":"VBScript_RegExp_55.SubMatches"},
		constants: {
		},
		enums: {
		},
	};
	return CACHE;
}
