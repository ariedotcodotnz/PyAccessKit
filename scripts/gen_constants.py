"""Generate ``src/pyaccesskit/_com/constants.py`` from the installed Access/DAO type libraries.

PyAccessKit never relies on makepy/``win32com.client.constants``. Instead, the handful of constants the
COM adapters need are generated from the *installed* type libraries (the real source of truth — they
contain values newer than the online documentation) and checked in. ``tests/integration`` contains a
drift test that re-renders this file and compares it with the checked-in copy.

Usage::

    uv run python scripts/gen_constants.py            # rewrite the module
    uv run python scripts/gen_constants.py --check    # exit 1 if the checked-in module is stale
"""

from __future__ import annotations

import argparse
import sys
import winreg
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes

OUTPUT = Path(__file__).resolve().parents[1] / "src" / "pyaccesskit" / "_com" / "constants.py"


@dataclass(frozen=True)
class TypeLibRef:
    """A registered type library and the enums/modules we want from it."""

    label: str
    guid: str
    major: int
    minor: int
    enums: tuple[str, ...]
    modules: tuple[str, ...] = ()


ACCESS = TypeLibRef(
    label="Microsoft Access 16.0 Object Library",
    guid="{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}",
    major=9,
    minor=0,
    enums=(
        "AcCloseSave",
        "AcControlType",
        "AcDefView",
        "AcFormOpenDataMode",
        "AcFormView",
        "AcModuleType",
        "AcNewDatabaseFormat",
        "AcObjectType",
        "AcQuitOption",
        "AcSection",
        "AcSysCmdAction",
        "AcView",
        "AcWindowMode",
    ),
)

DAO = TypeLibRef(
    label="Microsoft Office 16.0 Access database engine Object Library",
    guid="{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}",
    major=12,
    minor=0,
    enums=(
        "DataTypeEnum",
        "DatabaseTypeEnum",
        "FieldAttributeEnum",
        "QueryDefTypeEnum",
        "RecordsetOptionEnum",
        "RecordsetTypeEnum",
        "RelationAttributeEnum",
        "TableDefAttributeEnum",
    ),
    modules=("LanguageConstants",),
)

# Constants that live outside the Access/DAO type libraries (Office core, ADO, documented property
# values). They are rendered verbatim so the whole module stays generated and drift-checkable.
MANUAL_SECTION = '''

class MsoAutomationSecurity(IntEnum):
    """``Application.AutomationSecurity`` values (Office core type library)."""

    msoAutomationSecurityLow = 1
    msoAutomationSecurityByUI = 2
    msoAutomationSecurityForceDisable = 3


class AdoSchema(IntEnum):
    """ADO ``SchemaEnum`` values used for column introspection."""

    adSchemaColumns = 4
    adSchemaIndexes = 12


class FormScrollBars(IntEnum):
    """Values of the Access ``Form.ScrollBars`` property."""

    neither = 0
    horizontal = 1
    vertical = 2
    both = 3


class TextFormat(IntEnum):
    """Values of the Access ``TextFormat`` field property (Long Text)."""

    plain_text = 0
    rich_text = 1


# Well-known Win32/COM HRESULTs (as signed 32-bit integers, the way pywin32 reports them).
def _hresult(value: int) -> int:
    return value - 0x1_0000_0000 if value & 0x8000_0000 else value


DISP_E_EXCEPTION = _hresult(0x80020009)
REGDB_E_CLASSNOTREG = _hresult(0x80040154)
CO_E_SERVER_EXEC_FAILURE = _hresult(0x80080005)
RPC_E_CALL_REJECTED = _hresult(0x80010001)
RPC_E_SERVERCALL_RETRYLATER = _hresult(0x8001010A)
RPC_E_DISCONNECTED = _hresult(0x80010108)
RPC_E_SERVERFAULT = _hresult(0x80010105)
RPC_S_SERVER_UNAVAILABLE = _hresult(0x800706BA)
RPC_S_CALL_FAILED = _hresult(0x800706BE)
CO_E_OBJNOTCONNECTED = _hresult(0x800401FD)
'''


def _typelib_path(ref: TypeLibRef) -> str | None:
    """Resolve the on-disk path of a registered type library, trying win64 then win32."""
    key_base = rf"TypeLib\{ref.guid}\{ref.major:x}.{ref.minor:x}\0"
    for platform in ("win64", "win32"):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{key_base}\{platform}") as key:
                value, _ = winreg.QueryValueEx(key, "")
                return str(value)
        except OSError:
            continue
    return None


def _load(ref: TypeLibRef) -> Any:
    try:
        return pythoncom.LoadRegTypeLib(pywintypes.IID(ref.guid), ref.major, ref.minor, 0)
    except pythoncom.com_error:
        path = _typelib_path(ref)
        if path is None:
            raise SystemExit(
                f"{ref.label} ({ref.guid} v{ref.major}.{ref.minor}) is not registered"
            ) from None
        return pythoncom.LoadTypeLib(path)


def _render_typelib(ref: TypeLibRef) -> list[str]:
    tlb = _load(ref)
    found: dict[str, list[tuple[str, object]]] = {}
    wanted = set(ref.enums) | set(ref.modules)
    for index in range(tlb.GetTypeInfoCount()):
        name = tlb.GetDocumentation(index)[0]
        if name not in wanted:
            continue
        kind = tlb.GetTypeInfoType(index)
        if kind not in (pythoncom.TKIND_ENUM, pythoncom.TKIND_MODULE):
            continue
        info = tlb.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        members: list[tuple[str, object]] = []
        for var_index in range(attr.cVars):
            desc = info.GetVarDesc(var_index)
            members.append((info.GetNames(desc.memid)[0], desc.value))
        found[name] = members

    missing = sorted(wanted - found.keys())
    if missing:
        raise SystemExit(f"{ref.label}: missing {', '.join(missing)}")

    lines: list[str] = []
    for enum_name in ref.enums:
        lines += [
            "",
            "",
            f"class {enum_name}(IntEnum):",
            f'    """``{enum_name}`` ({ref.label})."""',
            "",
        ]
        lines += [f"    {member} = {int(value)}" for member, value in found[enum_name]]  # type: ignore[call-overload]
    for module_name in ref.modules:
        lines += ["", "", f"# {module_name} ({ref.label})"]
        for member, value in found[module_name]:
            lines.append(f"{member} = {value!r}")
    return lines


def render() -> str:
    """Render the full ``constants.py`` module text."""
    header = [
        '"""Access and DAO constants used by PyAccessKit\'s COM adapters (internal).',
        "",
        "Generated by ``scripts/gen_constants.py`` from the installed type libraries — do not edit by hand.",
        "",
        f"* {ACCESS.label} — ``{ACCESS.guid}`` v{ACCESS.major}.{ACCESS.minor}",
        f"* {DAO.label} — ``{DAO.guid}`` v{DAO.major}.{DAO.minor}",
        '"""',
        "",
        "# ruff: noqa: N815",
        "from __future__ import annotations",
        "",
        "from enum import IntEnum",
    ]
    body = _render_typelib(ACCESS) + _render_typelib(DAO)
    return "\n".join(header + body) + "\n\n\n" + MANUAL_SECTION.lstrip("\n")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check", action="store_true", help="fail if the checked-in module is stale"
    )
    args = parser.parse_args(argv)
    text = render()
    if args.check:
        current = OUTPUT.read_text(encoding="utf-8") if OUTPUT.exists() else ""
        if current != text:
            print(f"{OUTPUT} is stale; run scripts/gen_constants.py", file=sys.stderr)
            return 1
        print("constants are up to date")
        return 0
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(text, encoding="utf-8", newline="\n")
    print(f"wrote {OUTPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
