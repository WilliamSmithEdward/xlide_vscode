"""Unit tests for vba_io pure-logic helpers (no workbook I/O required)."""
from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from xlide.vba_io import _split_vba_source, _module_type


# ---------------------------------------------------------------------------
# _split_vba_source
# ---------------------------------------------------------------------------

class TestSplitVbaSource:
    def test_empty_source(self):
        header, body = _split_vba_source("")
        assert header == ""
        assert body == ""

    def test_plain_standard_module_has_no_header(self):
        src = "Option Explicit\r\nSub Hello()\r\nEnd Sub\r\n"
        header, body = _split_vba_source(src)
        assert header == ""
        assert "Sub Hello" in body

    def test_class_module_strips_version_block(self):
        src = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1  'True\r\n"
            "END\r\n"
            "Attribute VB_Name = \"Class1\"\r\n"
            "Attribute VB_GlobalNameSpace = False\r\n"
            "Option Explicit\r\n"
            "Sub Foo()\r\nEnd Sub\r\n"
        )
        header, body = _split_vba_source(src)
        assert "VERSION" in header
        assert "Attribute VB_Name" in header
        assert "Option Explicit" in body
        assert "VERSION" not in body

    def test_attribute_only_lines_go_into_header(self):
        src = "Attribute VB_Name = \"Mod\"\r\nSub Bar()\r\nEnd Sub\r\n"
        header, body = _split_vba_source(src)
        assert "Attribute VB_Name" in header
        assert "Sub Bar" in body
        assert "Attribute" not in body

    def test_join_round_trips(self):
        from xlide.vba_io import _join_vba_source
        src = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1\r\n"
            "END\r\n"
            "Attribute VB_Name = \"C\"\r\n"
            "Sub Foo()\r\nEnd Sub\r\n"
        )
        header, body = _split_vba_source(src)
        rejoined = _join_vba_source(header, body)
        # Both original sub-texts must survive the round-trip
        assert "VERSION" in rejoined
        assert "Sub Foo" in rejoined


# ---------------------------------------------------------------------------
# _module_type
# ---------------------------------------------------------------------------

class TestModuleType:
    def test_standard_no_attributes(self):
        result = _module_type("Module1", "Option Explicit\nSub Hello()\nEnd Sub\n")
        assert result == "standard"

    def test_userform_two_guids(self):
        src = (
            'Attribute VB_Base = '
            '"0{11111111-0000-0000-0000-000000000000};'
            '{22222222-0000-0000-0000-000000000000}"\n'
        )
        assert _module_type("UserForm1", src) == "userform"

    def test_workbook_document_clsid(self):
        src = 'Attribute VB_Base = "{00020819-0000-0000-C000-000000000046}"\n'
        assert _module_type("ThisWorkbook", src) == "document"

    def test_worksheet_document_clsid(self):
        src = 'Attribute VB_Base = "{00020820-0000-0000-C000-000000000046}"\n'
        assert _module_type("Sheet1", src) == "document"

    def test_chart_document_clsid(self):
        src = 'Attribute VB_Base = "{00020821-0000-0000-C000-000000000046}"\n'
        assert _module_type("Chart1", src) == "document"

    def test_thisworkbook_by_name(self):
        # No attributes — detected by name heuristic
        assert _module_type("ThisWorkbook", "Option Explicit\n") == "document"

    def test_sheet_by_name_pattern(self):
        assert _module_type("Sheet3", "Option Explicit\n") == "document"

    def test_predeclared_id_is_document(self):
        src = "Attribute VB_PredeclaredId = True\n"
        assert _module_type("Globals", src) == "document"

    def test_class_single_non_document_guid_returns_standard(self):
        # _module_type itself returns 'standard' for class modules;
        # list_modules upgrades to 'class' after checking VBAModuleKind.
        src = 'Attribute VB_Base = "{CC27B1A4-1234-1234-1234-000000000000}"\n'
        assert _module_type("MyClass", src) == "standard"


# ---------------------------------------------------------------------------
# allow_protected=True regression tests
# Verify every write path passes allow_protected=True to ExcelFile.save()
# so that VBA-project-password-protected workbooks are handled correctly.
# ---------------------------------------------------------------------------

def _make_mock_wb(module_source: str = "Sub Foo()\nEnd Sub\n") -> MagicMock:
    """Return a mock ExcelFile context manager whose module_names/get_module
    return predictable values."""
    mock_wb = MagicMock()
    mock_wb.__enter__ = MagicMock(return_value=mock_wb)
    mock_wb.__exit__ = MagicMock(return_value=False)
    mock_wb.module_names.return_value = ["ExistingModule"]
    mock_wb.get_module.return_value = module_source

    mock_project = MagicMock()
    mock_wb.vba_project.return_value = mock_project
    return mock_wb


_EXPECTED_SAVE_KWARGS = dict(allow_protected=True)


class TestAllowProtectedPassedToSave:
    """Every ExcelFile.save() call in vba_io must use allow_protected=True.
    The UserWarning for signature invalidation is captured via
    warnings.catch_warnings and surfaced in the signatureDropped response field
    rather than being silenced, so the TypeScript layer can notify the user."""

    def test_write_module_existing_passes_save_kwargs(self):
        mock_wb = _make_mock_wb()
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import write_module
            result = write_module(path="fake.xlsm", module="ExistingModule", source="Sub Bar()\nEnd Sub\n")
        mock_wb.save.assert_called_once_with(**_EXPECTED_SAVE_KWARGS)
        assert "signatureDropped" in result

    def test_write_module_new_passes_save_kwargs(self):
        mock_wb = _make_mock_wb()
        mock_wb.module_names.return_value = []  # new module path
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import write_module
            result = write_module(path="fake.xlsm", module="NewModule", source="Sub Baz()\nEnd Sub\n")
        mock_wb.save.assert_called_once_with(**_EXPECTED_SAVE_KWARGS)
        assert "signatureDropped" in result

    def test_rename_module_passes_save_kwargs(self):
        mock_wb = _make_mock_wb()
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import rename_module
            result = rename_module(path="fake.xlsm", module="OldName", newName="NewName")
        mock_wb.save.assert_called_once_with(**_EXPECTED_SAVE_KWARGS)
        assert "signatureDropped" in result

    def test_delete_module_passes_save_kwargs(self):
        mock_wb = _make_mock_wb()
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import delete_module
            result = delete_module(path="fake.xlsm", module="ExistingModule")
        mock_wb.save.assert_called_once_with(**_EXPECTED_SAVE_KWARGS)
        assert "signatureDropped" in result

    def test_signature_dropped_false_when_no_warning(self):
        mock_wb = _make_mock_wb()
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import write_module
            result = write_module(path="fake.xlsm", module="ExistingModule", source="Sub Bar()\nEnd Sub\n")
        assert result["signatureDropped"] is False

    def test_signature_dropped_true_when_user_warning_emitted(self):
        import warnings as _warnings

        mock_wb = _make_mock_wb()

        def _save_with_warning(**_kwargs: object) -> None:
            _warnings.warn("Dropped stale VBA digital signature", UserWarning, stacklevel=2)

        mock_wb.save.side_effect = _save_with_warning
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import write_module
            result = write_module(path="fake.xlsm", module="ExistingModule", source="Sub Bar()\nEnd Sub\n")
        assert result["signatureDropped"] is True


# ---------------------------------------------------------------------------
# get_protection_info
# ---------------------------------------------------------------------------

class TestGetProtectionInfo:
    """get_protection_info reports password-lock and signature presence using
    only public pyopenvba APIs (project.protection + detect_signature)."""

    def _run(self, *, has_protection: bool, has_password: bool, signed: bool):
        mock_wb = _make_mock_wb()
        mock_wb.vba_project_bytes.return_value = b"vbaproject"
        project = mock_wb.vba_project.return_value
        if has_protection:
            project.protection = MagicMock(has_password=has_password)
        else:
            project.protection = None

        sig = MagicMock(present=signed)
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb), \
                patch("pyopenvba.cfb.CFB", return_value=MagicMock()), \
                patch("pyopenvba.excel.detect_signature", return_value=sig):
            from xlide.vba_io import get_protection_info
            return get_protection_info(path="fake.xlsm")

    def test_unprotected_unsigned(self):
        result = self._run(has_protection=False, has_password=False, signed=False)
        assert result == {"isPasswordProtected": False, "isSigned": False}

    def test_password_protected(self):
        result = self._run(has_protection=True, has_password=True, signed=False)
        assert result["isPasswordProtected"] is True
        assert result["isSigned"] is False

    def test_protection_present_but_no_password(self):
        result = self._run(has_protection=True, has_password=False, signed=False)
        assert result["isPasswordProtected"] is False

    def test_signed(self):
        result = self._run(has_protection=False, has_password=False, signed=True)
        assert result["isSigned"] is True

    def test_protected_and_signed(self):
        result = self._run(has_protection=True, has_password=True, signed=True)
        assert result == {"isPasswordProtected": True, "isSigned": True}


# ---------------------------------------------------------------------------
# validate_workbook
# ---------------------------------------------------------------------------

class TestValidateWorkbook:
    def test_no_issues_returns_empty_list(self):
        mock_wb = _make_mock_wb()
        mock_wb.validate.return_value = []
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import validate_workbook
            result = validate_workbook(path="fake.xlsm")
        assert result == {"issues": []}

    def test_issues_are_passed_through(self):
        mock_wb = _make_mock_wb()
        mock_wb.validate.return_value = ["bad stream", "missing module"]
        with patch("xlide.vba_io.ExcelFile", return_value=mock_wb):
            from xlide.vba_io import validate_workbook
            result = validate_workbook(path="fake.xlsm")
        assert result == {"issues": ["bad stream", "missing module"]}


# ---------------------------------------------------------------------------
# create_workbook
# ---------------------------------------------------------------------------

class TestCreateWorkbook:
    def test_creates_and_closes(self):
        mock_wb = MagicMock()
        with patch("xlide.vba_io.ExcelFile") as mock_excel:
            mock_excel.create_new.return_value = mock_wb
            from xlide.vba_io import create_workbook
            result = create_workbook(path="new.xlsm")
        mock_excel.create_new.assert_called_once_with("new.xlsm")
        mock_wb.close.assert_called_once_with()
        assert result == {"ok": True, "path": "new.xlsm"}

