// Excel 16.0 measurements (Office Home 2024, 64-bit, 2026-10-15) of code a
// rule misjudged: each accepted case below compiled or ran where the named
// rule fired, and each rejected case is a control the rule still reports.
// The cases themselves are in syntax_corpus/oracle/vbe_oracle_cases.json.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../src/vbaProjectAnalysis';

interface OracleCase {
	id: string;
	source?: string;
	modules?: { name: string; type?: string; source: string }[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const byId = new Map(corpus.cases.map((oracleCase) => [oracleCase.id, oracleCase]));

/** Every code the case's modules report, as the extension analyzes an Excel workbook. */
function codesFor(id: string): string[] {
	const oracleCase = byId.get(id);
	if (!oracleCase) {
		throw new Error(`no oracle case ${id}`);
	}
	const modules: VbaProjectModuleInput[] = (oracleCase.modules ?? [{ name: 'Module1', type: 'standard', source: oracleCase.source ?? '' }])
		.map((mod) => ({ moduleName: mod.name, type: mod.type, source: mod.source }));
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: 'excel',
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.map((diagnostic) => diagnostic.code ?? ''));
}

const COMPILED_OR_RAN: ReadonlyArray<[string, readonly string[]]> = [
	['print_udt_len_argument_runtime', ['udt-value-mismatch']],
	['print_udt_lenb_argument_runtime', ['udt-value-mismatch']],
	['print_udt_len_second_argument_runtime', ['udt-value-mismatch']],
	['print_udt_user_function_argument_runtime', ['udt-value-mismatch']],
	['print_udt_varptr_argument_runtime', ['udt-value-mismatch']],
	['print_udt_len_plus_runtime', ['udt-value-mismatch']],
	['print_udt_grouped_len_runtime', ['udt-value-mismatch']],
	['print_udt_len_of_grouped_runtime', ['udt-value-mismatch']],
	['optional_collection_default_number_runtime', ['parameter-default-type-mismatch']],
	['optional_object_default_number_runtime', ['parameter-default-type-mismatch']],
	['optional_object_default_string_runtime', ['parameter-default-type-mismatch']],
	['optional_range_default_zero_runtime', ['parameter-default-type-mismatch']],
	['optional_collection_default_true_runtime', ['parameter-default-type-mismatch']],
	['optional_collection_default_date_runtime', ['parameter-default-type-mismatch']],
	['optional_collection_default_null_runtime', ['parameter-default-type-mismatch']],
	['optional_class_default_number_runtime', ['parameter-default-type-mismatch']],
	['type_named_range_shadows_excel_runtime', ['object-variable-not-set']],
	['type_named_worksheet_shadows_excel_runtime', ['object-variable-not-set']],
	['public_type_named_range_other_module_runtime', ['object-variable-not-set']],
	['enum_named_range_shadows_excel_runtime', ['object-variable-not-set']],
	['thisworkbook_type_set_activeworkbook_runtime', ['assignment-object-type-mismatch']],
	['thisworkbook_type_member_runtime', ['assignment-object-type-mismatch']],
	['thisworkbook_type_parameter_runtime', ['argument-type-mismatch']],
	['dim_shared_local_runtime', ['invalid-declaration-name', 'undeclared-variable', 'unexpected-declaration-token']],
	['dim_shared_two_names_runtime', ['invalid-declaration-name', 'unexpected-declaration-token']],
	['dim_shared_module_runtime', ['invalid-declaration-name', 'undeclared-variable', 'unexpected-declaration-token']],
	['private_shared_module_runtime', ['invalid-declaration-name', 'unexpected-declaration-token']],
	['dim_shared_array_runtime', ['invalid-declaration-name', 'undeclared-variable', 'unexpected-declaration-token']],
	['dim_shared_new_runtime', ['invalid-declaration-name', 'undeclared-variable', 'unexpected-declaration-token']],
	['public_shared_module_runtime', ['invalid-declaration-name', 'undeclared-variable', 'unexpected-declaration-token']],
	['global_shared_module_runtime', ['invalid-declaration-name', 'undeclared-variable', 'unexpected-declaration-token']],
	['dim_shared_dynamic_array_runtime', ['invalid-declaration-name', 'unexpected-declaration-token']],
	['erase_two_arrays_empty_parentheses_runtime', ['unallocated-dynamic-array-access']],
	['label_dollar_goto_runtime', ['undeclared-variable', 'undefined-label']],
	['label_dollar_goto_dollar_runtime', ['undeclared-variable']],
	['label_percent_runtime', ['undeclared-variable']],
	['label_hash_runtime', ['undeclared-variable']],
	['label_dollar_then_long_variable_runtime', ['type-suffix-mismatch', 'undeclared-variable']],
	['label_ampersand_runtime', ['undeclared-variable', 'undefined-label']],
	['label_bang_runtime', ['undeclared-variable', 'undefined-label']],
	['label_at_runtime', ['undeclared-variable', 'undefined-label']],
	['label_caret_runtime', ['invalid-expression-syntax', 'undeclared-variable', 'undefined-label']],
	['label_dollar_then_statement_runtime', ['undeclared-variable']],
	['label_dollar_resume_runtime', ['undeclared-variable', 'undefined-label']],
	['resume_double_zero_runtime', ['undefined-label']],
	['resume_zero_handler_runtime', ['undefined-label']],
	['err_unknown_member_assign_compile', ['member-not-found']],
	['err_unknown_member_read_compile', ['member-not-found']],
	['err_unknown_member_call_compile', ['member-not-found']],
	['me_let_document_module_compile', ['undeclared-variable']],
	['me_let_class_module_compile', ['undeclared-variable']],
	['me_let_keyword_class_compile', ['undeclared-variable']],
	['now_indexed_result_compile', ['argument-count']],
	['time_indexed_result_compile', ['argument-count']],
	['command_indexed_result_compile', ['argument-count']],
	['now_indexed_two_compile', ['argument-count']],
	['now_indexed_print_compile', ['argument-count']],
	['str_null_runtime', ['argument-type-mismatch']],
	['static_public_sub_runtime', ['unmatched-block-closer']],
	['static_friend_sub_runtime', ['unmatched-block-closer']],
	['byref_qualified_parenthesized_copy_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_byte_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_single_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_double_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_currency_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_string_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_date_runtime', ['byref-argument-type-mismatch']],
	['byref_qualified_parenthesized_boolean_runtime', ['byref-argument-type-mismatch']],
	['lset_string_number_runtime', ['string-arithmetic-coercion']],
	['rset_string_number_runtime', ['string-arithmetic-coercion']],
	['lset_string_long_variable_runtime', ['string-arithmetic-coercion']],
	['udt_power_compile', ['non-scalar-binary-operand', 'udt-value-mismatch']],
	['udt_power_right_compile', ['non-scalar-binary-operand', 'udt-value-mismatch']],
	['udt_not_runtime', ['udt-value-mismatch']],
	['array_power_runtime', ['non-scalar-binary-operand']],
	['udt_field_power_compile', ['non-scalar-binary-operand']],
	['label_line_number_and_dollar_runtime', ['undeclared-variable', 'undefined-label']],
	['document_type_sheet_worksheet_runtime', ['assignment-object-type-mismatch']],
	['str_null_into_variant_runtime', ['assignment-type-mismatch']],
	['optional_object_default_string_expression_runtime', ['parameter-default-not-constant', 'parameter-default-type-mismatch']],
	['line_number_then_label_caret_compile', ['invalid-expression-syntax', 'undefined-label']],
	['line_number_then_label_hash_compile', ['invalid-expression-syntax', 'undefined-label']],
	['label_caret_before_statement_runtime', ['invalid-expression-syntax', 'undefined-label']],
	['label_at_before_statement_runtime', ['invalid-expression-syntax', 'undefined-label']],
	['let_me_document_module_runtime', ['invalid-assignment-target']],
	['me_assign_document_module_runtime', ['invalid-assignment-target']],
	['project_time_variant_function_index_runtime', ['argument-count']],
	['static_private_function_compile', ['unmatched-block-closer']],
	['static_public_property_compile', ['unmatched-block-closer']],
];

const STILL_REPORTED: ReadonlyArray<[string, readonly string[]]> = [
	['print_udt_grouped_compile', ['udt-value-mismatch']],
	['print_udt_bare_compile', ['udt-value-mismatch']],
	['print_udt_as_array_index_compile', ['udt-value-mismatch']],
	['type_named_range_set_compile', ['set-requires-object']],
	['static_shared_compile', ['invalid-declaration-name', 'unexpected-declaration-token']],
	['dim_shared_as_type_compile', ['invalid-declaration-name']],
	['goto_zero_compile', ['undefined-label']],
	['resume_one_undefined_compile', ['undefined-label']],
	['resume_zero_without_error_runtime', ['resume-without-error']],
	['err_raise_too_many_arguments_compile', ['argument-count']],
	['set_me_class_compile', ['invalid-assignment-target']],
	['set_me_document_module_compile', ['invalid-assignment-target']],
	['timer_with_argument_compile', ['argument-count']],
	['erl_with_argument_compile', ['argument-count']],
	['doevents_with_argument_compile', ['argument-count']],
	['str_dollar_null_runtime', ['argument-type-mismatch']],
	['oct_dollar_null_runtime', ['argument-type-mismatch']],
	['byref_qualified_no_parentheses_compile', ['byref-argument-type-mismatch']],
	['byref_qualified_call_compile', ['byref-argument-type-mismatch']],
	['byref_qualified_function_compile', ['byref-argument-type-mismatch']],
	['udt_plus_compile', ['non-scalar-binary-operand', 'udt-value-mismatch']],
	['collection_power_compile', ['collection-operand']],
	['longlong_parenthesized_byref_long_compile', ['longlong-narrowing']],
	['longlong_qualified_parenthesized_byref_long_compile', ['longlong-narrowing']],
	['longlong_call_double_parentheses_compile', ['longlong-narrowing']],
	['longlong_named_parenthesized_compile', ['longlong-narrowing']],
	['longlong_qualified_byval_long_compile', ['longlong-narrowing']],
	['longlong_qualified_named_byval_long_compile', ['longlong-narrowing']],
	['longlong_named_byval_long_compile', ['longlong-narrowing']],
	['erase_empty_parentheses_long_compile', ['erase-requires-array']],
	['array_unary_minus_compile', ['non-scalar-binary-operand']],
	['document_type_sheet_workbook_runtime', ['assignment-object-type-mismatch']],
	['document_type_thisworkbook_application_runtime', ['assignment-object-type-mismatch']],
	['print_udt_array_index_in_call_compile', ['udt-value-mismatch']],
	['print_udt_array_index_single_line_if_compile', ['udt-value-mismatch']],
	['print_udt_field_array_index_compile', ['udt-value-mismatch']],
	['print_udt_collection_item_argument_compile', ['udt-value-mismatch']],
	['project_now_function_index_compile', ['argument-count']],
	['now_statement_parenthesized_compile', ['argument-count']],
	['time_dollar_argument_compile', ['argument-count']],
	['command_dollar_argument_compile', ['argument-count']],
	['str_null_into_string_runtime', ['assignment-type-mismatch']],
	['optional_object_default_array_call_compile', ['parameter-default-not-constant']],
	['set_me_standard_module_compile', ['me-outside-object-module']],
	['static_global_sub_header_compile', ['invalid-proc-header']],
];

describe('Excel-measured false positives', () => {
	it.each(COMPILED_OR_RAN)('%s: none of %j fires', (id, codes) => {
		const reported = codesFor(id);
		expect(codes.filter((code) => reported.includes(code))).toEqual([]);
	});
});

describe('Excel-measured controls the rules still report', () => {
	it.each(STILL_REPORTED)('%s reports %j', (id, codes) => {
		const reported = codesFor(id);
		expect(codes.filter((code) => !reported.includes(code))).toEqual([]);
	});
});
