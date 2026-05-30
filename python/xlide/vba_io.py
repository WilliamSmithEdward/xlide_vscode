"""VBA module read/write via pyOpenVBA."""
from __future__ import annotations

import re
import warnings
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


# Known document module CLSIDs (Excel Workbook, Worksheet, Chart).
# UserForms always carry TWO GUIDs in VB_Base — that pattern is the reliable
# discriminator and does not depend on any specific CLSID value.
_WORKBOOK_CLSID = "{00020819-0000-0000-C000-000000000046}"
_WORKSHEET_CLSID = "{00020820-0000-0000-C000-000000000046}"
_CHART_CLSID = "{00020821-0000-0000-C000-000000000046}"
_DOCUMENT_CLSIDS = (_WORKBOOK_CLSID, _WORKSHEET_CLSID, _CHART_CLSID)
_GUID_RE = re.compile(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}")


def _module_type(name: str, source: str) -> str:
    """Infer module type from source content and name.

    Returns one of: 'standard', 'class', 'document', 'userform'.
    """
    # Pull the VB_Base attribute value (if any).
    vb_base_match = re.search(
        r'^\s*Attribute\s+VB_Base\s*=\s*"([^"]*)"',
        source,
        re.MULTILINE | re.IGNORECASE,
    )
    vb_base = vb_base_match.group(1) if vb_base_match else ""

    if vb_base:
        # UserForms always have TWO GUIDs in VB_Base (type-lib + instance).
        # Class and document modules each have exactly one.
        if len(_GUID_RE.findall(vb_base)) >= 2:
            return "userform"
        if any(c in vb_base for c in _DOCUMENT_CLSIDS):
            return "document"

    if re.search(r"^\s*Attribute\s+VB_PredeclaredId\s*=\s*True", source, re.MULTILINE | re.IGNORECASE):
        return "document"
    # Well-known document-module names across common Excel locales.
    if name == "ThisWorkbook" or re.match(
        r"^(Sheet|Feuil|Hoja|Tabelle|Foglio|Planilha)\d*$", name, re.IGNORECASE
    ):
        return "document"
    return "standard"


def list_modules(*, path: str) -> list[dict[str, Any]]:
    """Return [{name, type}] for every VBA module in the workbook."""
    from pyopenvba import VBAModuleKind

    with ExcelFile(path) as wb:
        result = []
        for m in wb.vba_project().modules:
            if m.kind == VBAModuleKind.standard:
                mod_type = "standard"
            else:
                # VBAModuleKind.other covers both class and document modules.
                # Use source heuristics to distinguish them.
                mod_type = _module_type(m.name, m.source)
                if mod_type == "standard":
                    # .kind says it's not a standard module — treat as class.
                    mod_type = "class"
            result.append({"name": m.name, "type": mod_type})
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


def read_module(*, path: str, module: str, full: bool = False) -> dict[str, Any]:
    """Return {source} containing the module source.

    full=False (default) — returns only the user-visible body (no Attribute header).
    full=True            — returns the complete source including VERSION/Attribute headers,
                           suitable for exporting to files that need to round-trip accurately.
    """
    with ExcelFile(path) as wb:
        source = wb.get_module(module)
    if full:
        return {"source": source}
    _, body = _split_vba_source(source)
    return {"source": body}


def write_module(*, path: str, module: str, source: str, kind: str = "standard") -> dict[str, Any]:
    """Write source into a VBA module and save the workbook in place.

    ``source`` may be a bare visible body OR a full-source export (with
    VERSION/Attribute headers).  In either case the incoming header is stripped
    and replaced with the header already stored in the workbook, so the file
    round-trips cleanly regardless of what the caller provides.

    If the module does not yet exist it is created using ``kind``:
    - ``'standard'``: bas-style standard module
    - ``'class'``: cls-style class module (VB_PredeclaredId = False)
    """
    from pyopenvba import VBAModuleKind

    # Strip any incoming header so callers can pass full-export content safely.
    _, body = _split_vba_source(source)

    with ExcelFile(path) as wb:
        existing = wb.module_names()
        if module in existing:
            # Re-read the workbook's own header and re-attach it.
            current = wb.get_module(module)
            header, _ = _split_vba_source(current)
            full_source = _join_vba_source(header, body)
            wb.set_module(module, full_source)
        else:
            vba_kind = VBAModuleKind.other if kind == "class" else VBAModuleKind.standard
            project = wb.vba_project()
            project.add_module(module, body, kind=vba_kind)
        with warnings.catch_warnings(record=True) as _caught:
            warnings.simplefilter("always")
            wb.save(allow_protected=True)
    signature_dropped = any(issubclass(w.category, UserWarning) for w in _caught)
    return {"ok": True, "signatureDropped": signature_dropped}


def rename_module(*, path: str, module: str, newName: str) -> dict[str, Any]:
    """Rename a VBA module and save the workbook in place."""
    with ExcelFile(path) as wb:
        project = wb.vba_project()
        project.rename_module(module, newName)
        with warnings.catch_warnings(record=True) as _caught:
            warnings.simplefilter("always")
            wb.save(allow_protected=True)
    signature_dropped = any(issubclass(w.category, UserWarning) for w in _caught)
    return {"ok": True, "signatureDropped": signature_dropped}


def delete_module(*, path: str, module: str) -> dict[str, Any]:
    """Delete a VBA module and save the workbook in place."""
    with ExcelFile(path) as wb:
        project = wb.vba_project()
        project.delete_module(module)
        with warnings.catch_warnings(record=True) as _caught:
            warnings.simplefilter("always")
            wb.save(allow_protected=True)
    signature_dropped = any(issubclass(w.category, UserWarning) for w in _caught)
    return {"ok": True, "signatureDropped": signature_dropped}


def get_protection_info(*, path: str) -> dict[str, Any]:
    """Return {isPasswordProtected, isSigned} for the workbook's VBA project.

    - isPasswordProtected: the VBA project carries a password-lock record.
    - isSigned: the project has at least one digital-signature stream.

    Both flags are derived from public pyopenvba APIs only.
    """
    from pyopenvba.cfb import CFB
    from pyopenvba.excel import detect_signature

    with ExcelFile(path) as wb:
        project = wb.vba_project()
        is_protected = (
            project.protection is not None and project.protection.has_password
        )
        is_signed = detect_signature(CFB(wb.vba_project_bytes())).present
    return {
        "isPasswordProtected": bool(is_protected),
        "isSigned": bool(is_signed),
    }


def validate_workbook(*, path: str) -> dict[str, Any]:
    """Return {issues: [...]} listing cross-structure inconsistencies.

    An empty list means pyOpenVBA found no problems with the VBA project's
    internal structure.
    """
    with ExcelFile(path) as wb:
        issues = wb.validate()
    return {"issues": list(issues)}


def create_workbook(*, path: str) -> dict[str, Any]:
    """Create a new macro-enabled workbook with an empty VBA project.

    The workbook contains ThisWorkbook, Sheet1, and a bare Module1, built
    from pyOpenVBA's baked-in template so it opens cleanly in Excel.
    Supported extensions: .xlsm (default) and .xlsb.  Overwrites ``path``
    if it already exists.
    """
    wb = ExcelFile.create_new(path)
    wb.close()
    return {"ok": True, "path": path}

