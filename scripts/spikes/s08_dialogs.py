"""S8 — modal dialogs: can a watcher thread detect/dismiss them while the COM call blocks?

Also: Application.Run through raw IDispatch.Invoke (exact args) vs pywin32's typed wrapper.
"""

from __future__ import annotations

import contextlib
import threading
import time

import pythoncom
import pywintypes
import win32con
import win32gui
import win32process
from _harness import describe_com_error, fresh_dir, owned_access

acModule = 5
BM_CLICK = 0x00F5
PREFERENCE = ("cancel", "no", "ok", "close")


class Dismisser(threading.Thread):
    """Prototype watchdog: capture + dismiss visible windows of one PID (WM_CLOSE, then button click)."""

    def __init__(self, pid: int) -> None:
        super().__init__(daemon=True)
        self.pid = pid
        self.events: list[dict[str, object]] = []
        self.stop = threading.Event()

    def buttons(self, hwnd: int) -> list[tuple[int, str]]:
        out: list[tuple[int, str]] = []

        def cb(h: int, _: object) -> bool:
            if win32gui.GetClassName(h) == "Button":
                out.append((h, win32gui.GetWindowText(h)))
            return True

        with contextlib.suppress(Exception):
            win32gui.EnumChildWindows(hwnd, cb, None)
        return out

    def statics(self, hwnd: int) -> list[str]:
        out: list[str] = []

        def cb(h: int, _: object) -> bool:
            if win32gui.GetClassName(h) == "Static" and win32gui.GetWindowText(h):
                out.append(win32gui.GetWindowText(h))
            return True

        with contextlib.suppress(Exception):
            win32gui.EnumChildWindows(hwnd, cb, None)
        return out

    def run(self) -> None:
        while not self.stop.wait(0.2):
            found: list[int] = []

            def collect(h: int, _: object) -> bool:
                with contextlib.suppress(Exception):
                    if win32process.GetWindowThreadProcessId(h)[1] == self.pid and win32gui.IsWindowVisible(h):
                        found.append(h)
                return True

            with contextlib.suppress(Exception):
                win32gui.EnumWindows(collect, None)
            for hwnd in found:
                event = {
                    "t": round(time.perf_counter(), 2),
                    "class": win32gui.GetClassName(hwnd),
                    "title": win32gui.GetWindowText(hwnd),
                    "text": self.statics(hwnd),
                    "buttons": [b[1] for b in self.buttons(hwnd)],
                }
                if not any(e["title"] == event["title"] and e["text"] == event["text"] for e in self.events):
                    self.events.append(event)
                    print(f"  [dismisser] {event}")
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                time.sleep(0.5)
                if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                    btns = self.buttons(hwnd)
                    for wanted in PREFERENCE:
                        match = [h for h, text in btns if text.replace("&", "").strip().lower() == wanted]
                        if match:
                            print(f"  [dismisser] WM_CLOSE ignored; clicking {wanted!r}")
                            win32gui.PostMessage(match[0], BM_CLICK, 0, 0)
                            break


def invoke_exact(obj, name: str, *args):
    dispid = obj._oleobj_.GetIDsOfNames(name)
    return obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True, *args)


def main() -> None:
    work = fresh_dir("s08")
    with owned_access() as acc:
        app = acc.app
        app.AutomationSecurity = 1  # allow our own test code to run (Application.Run)
        app.NewCurrentDatabase(str(work / "dialogs.accdb"), 12)
        mod = work / "modRun.bas"
        mod.write_bytes(
            b"Option Compare Database\r\nPublic Function Add2(a As Long, b As Long) As Long\r\n    Add2 = a + b\r\nEnd Function\r\n"
            b"Public Function NoArgs() As String\r\n    NoArgs = \"ok\"\r\nEnd Function\r\n"
        )
        app.LoadFromText(acModule, "modRun", str(mod))

        dismisser = Dismisser(acc.pid)
        dismisser.start()
        try:
            t0 = time.perf_counter()
            r = app.Eval('MsgBox("Hello from spike", 0, "Spike OK box")')
            print(f"Eval(MsgBox OK) returned {r!r} after {time.perf_counter() - t0:.2f}s")
            t0 = time.perf_counter()
            r = app.Eval('MsgBox("Continue?", 4, "Spike YesNo box")')
            print(f"Eval(MsgBox YesNo) returned {r!r} after {time.perf_counter() - t0:.2f}s (6=Yes, 7=No)")
            t0 = time.perf_counter()
            try:
                app.DoCmd.OpenQuery("NoSuchQuery")
            except pywintypes.com_error as exc:
                print(f"OpenQuery missing -> {describe_com_error(exc)} after {time.perf_counter() - t0:.2f}s")
        finally:
            dismisser.stop.set()
            dismisser.join(3)
        print(f"dialog events: {len(dismisser.events)}")

        for label, fn in (
            ("app.Run('NoArgs') [pywin32 wrapper]", lambda: app.Run("NoArgs")),
            ("invoke_exact Run NoArgs", lambda: invoke_exact(app, "Run", "NoArgs")),
            ("invoke_exact Run Add2(2,3)", lambda: invoke_exact(app, "Run", "Add2", 2, 3)),
        ):
            try:
                print(f"{label} -> {fn()!r}")
            except pywintypes.com_error as exc:
                print(f"{label} FAILED: {describe_com_error(exc)}")
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
