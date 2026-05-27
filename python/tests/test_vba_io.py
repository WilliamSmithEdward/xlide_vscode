"""Unit tests for vba_io pure-logic helpers (no workbook I/O required)."""
from __future__ import annotations

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
