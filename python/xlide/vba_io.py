"""VBA module read/write via pyOpenVBA."""
from __future__ import annotations

import re
from typing import Any

from pyopenvba import ExcelFile

# ---------------------------------------------------------------------------
# Attribute-header handling
# ---------------------------------------------------------------------------
# The VBE hides all module-level header lines from the user.  These are:
#   - The VERSION / BEGIN / END block at the top of class modules
#   - All contiguous "Attribute VB_*" lines that follow
# We strip them before presenting source to the editor and re-attach the
# original header when writing back, so the workbook round-trips correctly.

_ATTR_LINE = re.compile(r"^Attribute\s+VB_", re.IGNORECASE)


def _split_vba_source(source: str) -> tuple[str, str]:
    """Return (hidden_header, visible_body).

    hidden_header — the VERSION/BEGIN/END block + Attribute VB_* lines.
    visible_body  — everything after, with leading blank lines stripped,
                    matching what the VBE shows.
    """
    lines = source.splitlines(keepends=True)
    i = 0

    # Class-module preamble: "VERSION x.x CLASS" / BEGIN / ... / END
    if lines and re.match(r"^VERSION\s+\d", lines[i], re.IGNORECASE):
        i += 1
        if i < len(lines) and lines[i].rstrip("\r\n").strip().upper() == "BEGIN":
            i += 1
            while i < len(lines) and lines[i].rstrip("\r\n").strip().upper() != "END":
                i += 1
            if i < len(lines):
                i += 1  # consume the END line

    # Module-level Attribute VB_* lines
    while i < len(lines) and _ATTR_LINE.match(lines[i]):
        i += 1

    header = "".join(lines[:i])
    body = "".join(lines[i:]).lstrip("\r\n")
    return header, body


def _join_vba_source(header: str, body: str) -> str:
    """Reconstruct the full source from a (header, body) pair."""
    if not header:
        return body
    # Ensure a single CRLF between the last header line and the body
    return header.rstrip("\r\n") + "\r\n" + body


# Matches Sub, Function, and Property procedures at any access level.
_PROC_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>\w+)\s*[(\r\n]",
    re.MULTILINE | re.IGNORECASE,
)


def _module_type(name: str, source: str) -> str:
    """Infer module type from source content and name."""
    stripped = source.lstrip()
    if stripped.lower().startswith("version 1.0 class"):
        return "class"
    if "Attribute VB_PredeclaredId = True" in source:
        return "document"
    # Well-known document-module names across common Excel locales.
    if name == "ThisWorkbook" or re.match(
        r"^(Sheet|Feuil|Hoja|Tabelle|Foglio|Planilha)\d*$", name, re.IGNORECASE
    ):
        return "document"
    return "standard"


def list_modules(*, path: str) -> list[dict[str, Any]]:
    """Return [{name, type}] for every VBA module in the workbook."""
    with ExcelFile(path) as wb:
        names: list[str] = wb.module_names()
        result = []
        for name in names:
            source = wb.get_module(name)
            result.append({"name": name, "type": _module_type(name, source)})
        return result


def list_subs(*, path: str, module: str) -> list[dict[str, Any]]:
    """Return [{name, kind, line}] for every procedure in a module.

    Line numbers are 1-based and relative to the visible body (no header),
    matching what the editor shows.
    """
    with ExcelFile(path) as wb:
        source = wb.get_module(module)

    _, body = _split_vba_source(source)
    subs: list[dict[str, Any]] = []
    for match in _PROC_RE.finditer(body):
        line_num = body[: match.start()].count("\n") + 1
        subs.append(
            {
                "name": match.group("name"),
                "kind": match.group("kind").strip(),
                "line": line_num,
            }
        )
    return subs


def read_module(*, path: str, module: str) -> dict[str, Any]:
    """Return {source} containing only the user-visible body (no Attribute header)."""
    with ExcelFile(path) as wb:
        source = wb.get_module(module)
    _, body = _split_vba_source(source)
    return {"source": body}


def write_module(*, path: str, module: str, source: str) -> dict[str, Any]:
    """Write source into a VBA module and save the workbook in place.

    If the module does not yet exist it is created as a standard module.
    """
    from pyopenvba import VBAModuleKind

    with ExcelFile(path) as wb:
        existing = wb.module_names()
        if module in existing:
            # Re-read the current header and re-attach it so the workbook
            # round-trips cleanly even though the editor only shows the body.
            current = wb.get_module(module)
            header, _ = _split_vba_source(current)
            full_source = _join_vba_source(header, source)
            wb.set_module(module, full_source)
        else:
            project = wb.vba_project()
            project.add_module(module, source, kind=VBAModuleKind.standard)
        wb.save()
    return {"ok": True}


def rename_module(*, path: str, module: str, newName: str) -> dict[str, Any]:
    """Rename a VBA module and save the workbook in place."""
    with ExcelFile(path) as wb:
        project = wb.vba_project()
        project.rename_module(module, newName)
        wb.save()
    return {"ok": True}


def delete_module(*, path: str, module: str) -> dict[str, Any]:
    """Delete a VBA module and save the workbook in place."""
    with ExcelFile(path) as wb:
        project = wb.vba_project()
        project.delete_module(module)
        wb.save()
    return {"ok": True}
