"""Real-workbook coverage for excel_io workbook summaries.

These exercise get_workbook_info/list_sheets against the checked-in test
workbook because openpyxl's ReadOnlyWorksheet API differs from the writable
worksheet API (no `dimensions` property since openpyxl 3.1) — a difference
mocks cannot catch.
"""
import os
import re
import shutil

import pytest

from xlide.excel_io import get_workbook_info, list_sheets

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_TEST_WORKBOOK = os.path.join(_REPO_ROOT, "excel_test_workbook", "fullBuild.xlsm")

_DIMENSIONS_RE = re.compile(r"^$|^[A-Z]+\d+(:[A-Z]+\d+)?$")


@pytest.fixture()
def workbook_copy(tmp_path):
    if not os.path.exists(_TEST_WORKBOOK):
        pytest.skip("excel_test_workbook/fullBuild.xlsm not present")
    dest = tmp_path / "fullBuild.xlsm"
    shutil.copyfile(_TEST_WORKBOOK, dest)
    return str(dest)


def test_list_sheets_returns_sheet_summaries(workbook_copy):
    result = list_sheets(path=workbook_copy)
    sheets = result["sheets"]
    assert sheets, "expected at least one worksheet"
    for sheet in sheets:
        assert sheet["name"]
        assert _DIMENSIONS_RE.match(sheet["dimensions"]), sheet


def test_get_workbook_info_combines_sheets_modules_and_names(workbook_copy):
    info = get_workbook_info(path=workbook_copy)
    assert info["sheets"], "expected at least one worksheet summary"
    for sheet in info["sheets"]:
        assert _DIMENSIONS_RE.match(sheet["dimensions"]), sheet
    assert isinstance(info["namedRanges"], list)
    assert "modules" in info
