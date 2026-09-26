# pyright: basic
"""Environment probing: which engines can this Python process use, and why not?

Registry checks use the standard library (``winreg``); the in-process DAO check actually loads
``DAO.DBEngine.120`` (cheap, no process is started). Results are cached per process.
"""

from __future__ import annotations

import functools
import os
import struct
import winreg
from dataclasses import dataclass

import pywintypes

from pyaccesskit._com.dispatch import create_inproc

__all__ = [
    "ACE_OLEDB_PROGIDS",
    "DAO_PROGID",
    "AccessFacts",
    "ClickToRunFacts",
    "ProbeResult",
    "access_facts",
    "access_registered",
    "click_to_run",
    "dao_registration",
    "file_version",
    "inproc_dao",
    "inproc_registration",
    "pe_bits",
    "python_bits",
]

DAO_PROGID = "DAO.DBEngine.120"
ACE_OLEDB_PROGIDS = ("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")
_APP_PATHS = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE"
_C2R_CONFIGURATION = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
_PE_MACHINE_BITS = {0x014C: 32, 0x8664: 64, 0xAA64: 64}


@dataclass(frozen=True)
class ProbeResult:
    """Whether something is usable, with a human explanation when it is not."""

    available: bool
    reason: str = ""


def python_bits() -> int:
    """Bitness of this Python interpreter (32 or 64)."""
    return struct.calcsize("P") * 8


def _clsid(progid: str) -> str | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CLSID") as key:
            return str(winreg.QueryValueEx(key, "")[0])
    except OSError:
        return None


def _inproc_server(clsid: str, view: int) -> str | None:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\InprocServer32", 0, winreg.KEY_READ | view
        ) as key:
            value = str(winreg.QueryValueEx(key, "")[0])
            return value or None
    except OSError:
        return None


def inproc_registration(progid: str) -> dict[int, str | None]:
    """In-process server path registered for ``progid``, for 32- and 64-bit processes."""
    clsid = _clsid(progid)
    if clsid is None:
        return {32: None, 64: None}
    return {
        32: _inproc_server(clsid, winreg.KEY_WOW64_32KEY),
        64: _inproc_server(clsid, winreg.KEY_WOW64_64KEY),
    }


def dao_registration() -> dict[int, str | None]:
    """Native in-process DAO server path registered for 32- and 64-bit processes."""
    return inproc_registration(DAO_PROGID)


def access_registered(progid: str = "Access.Application") -> bool:
    """Whether ``Access.Application`` (or the given ProgID) is registered."""
    return _clsid(progid) is not None


@functools.cache
def inproc_dao() -> ProbeResult:
    """Whether this Python process can load DAO in-process (cached)."""
    try:
        engine = create_inproc(DAO_PROGID)
    except pywintypes.com_error:
        registered = dao_registration()
        bits = python_bits()
        other = 64 if bits == 32 else 32
        if registered[bits]:
            reason = f"DAO is registered for {bits}-bit processes but could not be loaded"
        elif registered[other]:
            reason = (
                f"DAO (Access database engine) is installed for {other}-bit programs only, but this Python is "
                f"{bits}-bit. Use a {other}-bit Python for in-process DAO, or engine='access'."
            )
        else:
            reason = "the Access database engine (DAO.DBEngine.120) is not installed"
        return ProbeResult(False, reason)
    del engine
    return ProbeResult(True)


# ------------------------------------------------------------------------------ installation facts
@dataclass(frozen=True)
class AccessFacts:
    """What the registry says about Microsoft Access (nothing is started)."""

    progid: str
    registered: bool
    current_version: str | None
    executable: str | None
    version: str | None
    bits: int | None


@dataclass(frozen=True)
class ClickToRunFacts:
    """Microsoft 365 / Click-to-Run installation details."""

    platform: str | None
    version: str | None
    products: tuple[str, ...]


def _read_value(root: int, path: str, name: str, view: int) -> str | None:
    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_READ | view) as key:
            value = winreg.QueryValueEx(key, name)[0]
    except OSError:
        return None
    return str(value) if value not in (None, "") else None


def _executable_from_command(command: str) -> str | None:
    command = command.strip()
    if command.startswith('"'):
        end = command.find('"', 1)
        return command[1:end] if end > 1 else None
    lowered = command.lower()
    end = lowered.find(".exe")
    return command[: end + 4] if end >= 0 else command or None


def _local_server(clsid: str) -> str | None:
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        command = _read_value(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32", "", view)
        if command:
            return _executable_from_command(command)
    return None


def pe_bits(path: str) -> int | None:
    """Bitness of a Windows executable, read from its PE header (``None`` if unreadable)."""
    try:
        with open(path, "rb") as handle:  # noqa: PTH123 - plain binary read of a header
            header = handle.read(64)
            if header[:2] != b"MZ" or len(header) < 64:
                return None
            handle.seek(int.from_bytes(header[0x3C:0x40], "little"))
            signature = handle.read(6)
    except OSError:
        return None
    if signature[:4] != b"PE\0\0":
        return None
    return _PE_MACHINE_BITS.get(int.from_bytes(signature[4:6], "little"))


def file_version(path: str) -> str | None:
    """The file version resource of an executable (e.g. ``16.0.19127.20264``)."""
    try:
        import win32api

        info = win32api.GetFileVersionInfo(path, "\\")
    except Exception:
        return None
    ms, ls = int(info["FileVersionMS"]), int(info["FileVersionLS"])
    return f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"


def access_facts(progid: str = "Access.Application") -> AccessFacts:
    """Locate the Access executable registered for automation (registry only)."""
    clsid = _clsid(progid)
    current = _read_value(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CurVer", "", winreg.KEY_WOW64_64KEY)
    executable = _local_server(clsid) if clsid else None
    if executable is None:
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            executable = _read_value(winreg.HKEY_LOCAL_MACHINE, _APP_PATHS, "", view)
            if executable:
                break
    if executable and not os.path.isfile(executable):  # noqa: PTH113 - registry path strings
        executable = None
    return AccessFacts(
        progid=progid,
        registered=clsid is not None,
        current_version=current,
        executable=executable,
        version=file_version(executable) if executable else None,
        bits=pe_bits(executable) if executable else None,
    )


def click_to_run() -> ClickToRunFacts | None:
    """Click-to-Run configuration, or ``None`` for MSI installations / no Office."""
    view = winreg.KEY_WOW64_64KEY
    platform = _read_value(winreg.HKEY_LOCAL_MACHINE, _C2R_CONFIGURATION, "Platform", view)
    version = _read_value(winreg.HKEY_LOCAL_MACHINE, _C2R_CONFIGURATION, "VersionToReport", view)
    products = _read_value(winreg.HKEY_LOCAL_MACHINE, _C2R_CONFIGURATION, "ProductReleaseIds", view)
    if platform is None and version is None and products is None:
        return None
    return ClickToRunFacts(
        platform=platform,
        version=version,
        products=tuple(p for p in (products or "").split(",") if p),
    )
