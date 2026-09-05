"""Measure the value Access gives every design property of every control type.

An Access design stores a property only where it differs from what the object
would have anyway: a command button whose font weight is the ordinary one
carries no `FontWeight` record, and a reader looking at the control alone finds
nothing to show. Two layers fill that in - the design's own control-defaults
object for the type, and, under that, the value Access gives a control it has
just created. This measures the second.

Run it with Access installed; it writes `src/vba/access/accessControlDefaults.ts`.
Nothing here is inferred: every value comes from a control Access made, read
back through the same property names the design schema uses.

    python scripts/measure-access-control-defaults.py
"""
from __future__ import annotations

import os
import pathlib
import re
import sys
import tempfile

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover - the script is Windows-only by nature
    sys.exit("This needs pywin32 and a local Access install.")

ROOT = pathlib.Path(__file__).resolve().parent.parent
TABLE = ROOT / "src" / "vba" / "access" / "accessDesignTable.ts"
OUT = ROOT / "src" / "vba" / "access" / "accessControlDefaults.ts"

ACC_FORM = 2
ACC_REPORT = 3
ACC_DESIGN = 1

# Geometry is the control's own, and identity is not a default.
SKIP = {
    "Name", "Left", "Top", "Width", "Height", "GUID",
    "LayoutCachedLeft", "LayoutCachedTop", "LayoutCachedWidth", "LayoutCachedHeight",
    "OverlapFlags", "IMESentenceMode",
}

# Where the design schema's name for a property is not the object model's.
ALIASES = {"OldBorderStyle": "BorderStyle"}

# What the Office theme decides, which is machine-specific and belongs to the
# design rather than to a shipped table. Access writes these onto the design's
# own control-defaults object for each type it uses, so a reader already has
# them from there; measuring them here would ship one machine's theme as
# everybody's default and draw a colour the control does not have.
THEMED_SUFFIXES = ("Color", "ThemeColorIndex", "ThemeFontIndex", "Tint", "Shade")
THEMED = {
    "FontName", "FontSize", "FontWeight", "Gradient", "Shape", "UseTheme",
    "QuickStyleMask", "Bevel", "SoftEdges",
}


def themed(name: str) -> bool:
    return name in THEMED or name.endswith(THEMED_SUFFIXES)

SYSTEM_COLOR_NAMES = [
    "ScrollBars", "Desktop", "ActiveTitleBar", "InactiveTitleBar", "MenuBar",
    "WindowBackground", "WindowFrame", "MenuText", "WindowText", "TitleBarText",
    "ActiveBorder", "InactiveBorder", "ApplicationWorkspace", "Highlight",
    "HighlightText", "ButtonFace", "ButtonShadow", "GrayText", "ButtonText",
    "InactiveTitleBarText", "ButtonHighlight", "ButtonDarkShadow", "ButtonLight",
    "InfoText", "InfoBackground", "HotTracking", "GradientActiveTitleBar",
    "GradientInactiveTitleBar", "MenuHighlight", "MenuBackground",
]

SLOT = re.compile(r"\['([A-Za-z0-9_]+)', \[(\d+), (\d+), (\d+), (\d+)(?:, (\d+))?\]\]")
SCHEMA = re.compile(r"\t\t\['([A-Za-z0-9_]+)', new Map\(\[([\s\S]*?)\]\)\],\n")


def read_schemas() -> dict[str, dict[str, tuple[int, int, int, int, int]]]:
    """Every object type's properties, from the generated slot table."""
    text = TABLE.read_text(encoding="utf-8")
    body = text[text.index("export const PROPERTY_SLOTS"):text.index("export const COMPANIONS")]
    out: dict[str, dict[str, tuple[int, int, int, int, int]]] = {}
    for schema in SCHEMA.finditer(body):
        props: dict[str, tuple[int, int, int, int, int]] = {}
        for slot in SLOT.finditer(schema.group(2)):
            props[slot.group(1)] = (
                int(slot.group(2)), int(slot.group(3)), int(slot.group(4)),
                int(slot.group(5)), int(slot.group(6) or 0),
            )
        out[schema.group(1)] = props
    return out


def read_control_types() -> dict[str, int]:
    text = TABLE.read_text(encoding="utf-8")
    body = text[text.index("export const CONTROL_TYPES"):text.index("export const PROPERTY_CODES")]
    return {name: int(code) for code, name in re.findall(r"\[(\d+), '([A-Za-z]+)'\]", body)}


def ole_color(value: int) -> str:
    raw = value & 0xFFFFFFFF
    if (raw >> 24) & 0xFF == 0x80 and (raw & 0xFFFFFF) < len(SYSTEM_COLOR_NAMES):
        return SYSTEM_COLOR_NAMES[raw & 0xFFFFFF]
    return "#%02x%02x%02x" % (raw & 0xFF, (raw >> 8) & 0xFF, (raw >> 16) & 0xFF)


def printed(name: str, slot: tuple[int, int, int, int, int], value: object) -> str | None:
    """The value as the markup and the property pane spell it."""
    value_type, length = slot[2], slot[4]
    if value is None:
        return None
    if value_type in (10, 12):
        return str(value)
    if value_type == 1:
        return "True" if value else "False"
    if value_type == 4 and length == 4 and name.endswith("Color"):
        try:
            return ole_color(int(value))
        except (TypeError, ValueError):
            return None
    if value_type in (6, 8):
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return str(int(number)) if number.is_integer() else repr(number)
    try:
        return str(int(value))
    except (TypeError, ValueError):
        return None


def section_of(design, index: int):
    """One band of the design.

    `Form.Section(acDetail)` is a property that takes an argument, which late
    binding turns into a method call Access refuses. Invoking it as a property
    get by dispatch id is the same call VBA makes.
    """
    try:
        ole = design._oleobj_
        raw = ole.Invoke(
            ole.GetIDsOfNames("Section"), 0, pythoncom.DISPATCH_PROPERTYGET, 1, index,
        )
        return win32com.client.Dispatch(raw)
    except Exception:  # noqa: BLE001
        return None


def properties_of(obj) -> dict[str, object]:
    """Every property the object exposes, by name, skipping the unreadable."""
    out: dict[str, object] = {}
    try:
        count = obj.Properties.Count
    except Exception:  # noqa: BLE001
        return out
    for i in range(count):
        try:
            prop = obj.Properties(i)
            out[prop.Name] = prop.Value
        except Exception:  # noqa: BLE001
            continue
    return out


def measure() -> dict[str, dict[str, str]]:
    schemas = read_schemas()
    codes = read_control_types()
    work = tempfile.mkdtemp(prefix="xlide_defaults_")
    db = os.path.join(work, "defaults.accdb")

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    measured: dict[str, dict[str, str]] = {}
    try:
        app.NewCurrentDatabase(db)
        for kind, opener in (("form", app.CreateForm), ("report", app.CreateReport)):
            design = opener()
            design_name = design.Name
            schema_of_design = schemas.get("_Design", {})
            raw = properties_of(design)
            record(measured, "Form" if kind == "form" else "Report", schema_of_design, raw)
            # The sections the design has. Access numbers them acDetail,
            # acHeader, acFooter, acPageHeader, acPageFooter.
            for index, schema_name in ((0, "Detail"), (1, "HeaderSection"), (2, "FooterSection"),
                                       (3, "PageHeaderSection"), (4, "PageFooterSection")):
                if schema_name not in schemas:
                    continue
                section = section_of(design, index)
                if section is None:
                    print(f"  {schema_name}: the design has none", file=sys.stderr)
                    continue
                record(measured, schema_name, schemas[schema_name], properties_of(section))
            if kind == "form":
                made: dict[str, str] = {}
                order = sorted(
                    codes.items(),
                    key=lambda pair: (pair[0] in ("Page", "NavigationButton"), pair[0]),
                )
                for schema_name, code in order:
                    if schema_name not in schemas or code >= 152:
                        continue
                    # A page belongs to a tab control and a navigation button
                    # to a navigation control; Access refuses either on its own.
                    parent = ""
                    if schema_name == "Page":
                        parent = made.get("Tab", "")
                    elif schema_name == "NavigationButton":
                        parent = made.get("NavigationControl", "")
                    try:
                        control = app.CreateControl(
                            design_name, code, 0, parent, "", 100, 100, 1440, 300,
                        )
                    except Exception as err:  # noqa: BLE001
                        print(f"  {schema_name}: {err}", file=sys.stderr)
                        continue
                    made[schema_name] = control.Name
                    record(measured, schema_name, schemas[schema_name], properties_of(control))
            app.DoCmd.Close(ACC_FORM if kind == "form" else ACC_REPORT, design_name, 2)
        app.CloseCurrentDatabase()
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
    return measured


def record(
    measured: dict[str, dict[str, str]],
    schema_name: str,
    slots: dict[str, tuple[int, int, int, int, int]],
    raw: dict[str, object],
) -> None:
    by_lower = {name.lower(): value for name, value in raw.items()}
    out = measured.setdefault(schema_name, {})
    for name, slot in slots.items():
        if name in SKIP or name.startswith("Unidentified") or themed(name):
            continue
        source = ALIASES.get(name, name)
        if source.lower() not in by_lower:
            continue
        text = printed(name, slot, by_lower[source.lower()])
        if text is not None:
            out[name] = text
    print(f"  {schema_name}: {len(out)} of {len(slots)}", file=sys.stderr)


def emit(measured: dict[str, dict[str, str]]) -> str:
    lines = [
        "// GENERATED FILE - do not edit by hand.",
        "//",
        "// The value Access gives each design property of a control it has just",
        "// created, read back through the object model with",
        "// `scripts/measure-access-control-defaults.py`. Regenerate from that.",
        "//",
        "// A design stores a property only where it differs from what the object",
        "// would have anyway, so a reader needs this to show a complete property",
        "// sheet: the control's own record first, then the design's",
        "// control-defaults object for its type, then this.",
        "//",
        "// What the Office theme decides - every colour, font, tint and shade - is",
        "// left out. Access writes those onto the design's own control-defaults",
        "// object, where they are the theme that database actually uses; a table",
        "// here would ship the measuring machine's theme as everybody's.",
        "",
        "/** Access's own value for a property, by object type and property name. */",
        "export const ACCESS_CONTROL_DEFAULTS:"
        " ReadonlyMap<string, ReadonlyMap<string, string>> =",
        "\tnew Map([",
    ]
    for schema_name in sorted(measured):
        props = measured[schema_name]
        if not props:
            continue
        lines.append(f"\t\t['{schema_name}', new Map([")
        for name in sorted(props):
            value = props[name].replace("\\", "\\\\").replace("'", "\\'")
            lines.append(f"\t\t\t['{name}', '{value}'],")
        lines.append("\t\t])],")
    lines.append("\t]);")
    return "\n".join(lines) + "\n"


if __name__ == "__main__":
    print("measuring...", file=sys.stderr)
    OUT.write_text(emit(measure()), encoding="utf-8")
    print(f"wrote {OUT}", file=sys.stderr)
