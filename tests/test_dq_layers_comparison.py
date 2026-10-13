"""Tests for the bronze/silver normalisation.

The comparison is only useful if it is wrong in neither direction: it must
not report a dialect difference as a discrepancy, and it must not quietly
swallow a real one. The cases below are drawn from the sprint 1-3 exports.
"""

from __future__ import annotations

import pytest

from dge.ded.scripts.dq_layers_comparison import (
    is_sentinel,
    normalise_for,
    normalise_properties,
    normalise_sql,
)


def same(column: str, bronze, silver) -> bool:
    return normalise_for(column, bronze) == normalise_for(column, silver)


# --------------------------------------------------------------------------
# differences the comparison should absorb
# --------------------------------------------------------------------------


def test_identifier_case_is_dialect():
    assert same("column", "B1_PER_ID1,SERV_PROV_CODE", "b1_per_id1,serv_prov_code")


def test_spacing_inside_a_call_is_dialect():
    """The pair that prompted this: `isNull(X)` against `isNull( x )`."""
    assert same(
        "rule_expression_logic",
        "isNull(MACHINE_ADDRESS) OR MACHINE_ADDRESS RLIKE '^[^ ]*( [^ ]+)*$'",
        "isNull( machine_address ) OR machine_address RLIKE '^[^ ]*( [^ ]+)*$'",
    )


def test_templated_braces_and_double_equals_are_dialect():
    assert same("filter_Expression_logic", "{REC_STATUS} == 'A'", "rec_status = 'A'")


def test_backticked_identifiers_are_dialect():
    assert same("rule_expression_logic", "CAST(`hs-code` AS STRING)", "cast(hs-code as string)")


def test_type_properties_fold_identifier_case_and_column_order():
    bronze = (
        "{'condition': \"isNull(MACHINE_ADDRESS)\", "
        "'columns': [{'value': 'MACHINE_ADDRESS', 'type': 'Column'}, "
        "{'value': 'REC_STATUS', 'type': 'Column'}], "
        "'emptyCriteria': None, 'filterCriteria': \"REC_STATUS = 'A'\", "
        "'ruleDialect': 'SparkSQL'}"
    )
    silver = (
        "{'condition': \"isNull( machine_address )\", "
        "'columns': [{'type': 'Column', 'value': 'rec_status'}, "
        "{'type': 'Column', 'value': 'machine_address'}], "
        "'filterCriteria': \"rec_status = 'A'\", "
        "'ruleDialect': 'SparkSQL'}"
    )
    assert same("typeProperties", bronze, silver)


def test_absent_key_matches_an_explicit_none():
    assert same("typeProperties", "{'a': 1}", "{'a': 1, 'emptyCriteria': None}")


def test_dataset_name_inside_properties_is_layer_specific():
    assert same(
        "typeProperties",
        "{'referencedDatasetAlias': 'virtual_commercial_activities', 'condition': 'x > 1'}",
        "{'referencedDatasetAlias': 'commercial_activities', 'condition': 'x > 1'}",
    )


def test_blank_and_none_agree():
    assert same("status", None, "")
    assert same("score", 1, 1.0)


# --------------------------------------------------------------------------
# differences the comparison must keep reporting
# --------------------------------------------------------------------------


def test_spacing_inside_a_string_literal_is_meaningful():
    """The regex depends on its spaces, so these are genuinely different."""
    assert not same(
        "rule_expression_logic",
        "x RLIKE '^[^ ]*( [^ ]+)*$'",
        "x RLIKE '^[^]*([^]+)*$'",
    )


def test_literal_case_is_meaningful():
    assert not same("filter_Expression_logic", "rec_status = 'A'", "rec_status = 'a'")


def test_different_logic_is_reported():
    assert not same(
        "rule_expression_logic",
        "UPPER(x) NOT RLIKE '^[ ]*(NA|N/A)'",
        "UPPER(TRIM(x)) NOT IN ('NA','N/A')",
    )


def test_escaping_difference_is_reported():
    assert not same("rule_expression_logic", "x RLIKE '\\pCntrl'", "x RLIKE '\\\\pCntrl'")


def test_a_different_rule_shape_is_reported():
    """Bronze implements the check as an expression, silver as a pattern."""
    assert not same(
        "typeProperties",
        "{'columns': [{'value': 'A', 'type': 'Column'}], 'condition': 'regexMatch({A}, \"^x$\")'}",
        "{'column': {'value': 'a', 'type': 'Column'}, 'pattern': '^x$'}",
    )


def test_an_extra_column_is_reported():
    assert not same(
        "typeProperties",
        "{'columns': [{'value': 'A', 'type': 'Column'}, {'value': 'B', 'type': 'Column'}]}",
        "{'columns': [{'value': 'a', 'type': 'Column'}]}",
    )


def test_a_populated_criterion_does_not_match_a_missing_one():
    assert not same(
        "typeProperties",
        "{'emptyCriteria': None}",
        "{'emptyCriteria': 'x IS NULL'}",
    )


def test_operators_are_not_discarded_wholesale():
    assert not same("rule_expression_logic", "a = b", "a <= b")
    assert not same("rule_expression_logic", "a > 1", "a < 1")


# --------------------------------------------------------------------------
# robustness
# --------------------------------------------------------------------------


def test_unparseable_properties_fall_back_to_the_tokeniser():
    assert normalise_properties("{not a dict")[0] == "unparsed"
    assert same("typeProperties", "{not a dict", "{NOT A DICT")


def test_normalisers_accept_none_and_odd_types():
    assert normalise_sql(None) == ()
    assert normalise_properties(None) == ()
    assert normalise_for("typeProperties", 3) == normalise_for("typeProperties", 3)


def test_an_unbalanced_quote_does_not_hang_or_crash():
    assert normalise_sql("x = 'unterminated") != ()


@pytest.mark.parametrize("value", ["N/A", "n/a", "N?A", "NA", "", "  ", "null"])
def test_sentinels(value):
    assert is_sentinel(value)


@pytest.mark.parametrize("value", ["DQO-475", "b5b2e534-8e1e-4619-83cb-b88c11a18c92"])
def test_real_identifiers_are_not_sentinels(value):
    assert not is_sentinel(value)


def test_column_list_order_is_not_meaningful():
    assert same(
        "column",
        "B1_PER_ID1,B1_PER_ID2,ROW_INDEX,SERV_PROV_CODE",
        "b1_per_id1,b1_per_id2,serv_prov_code,row_index",
    )


def test_a_missing_column_in_the_list_is_reported():
    assert not same("column", "A,B,C", "a,b")
