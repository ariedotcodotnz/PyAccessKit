"""S1 — process ownership: new instance per CoCreateInstanceEx, PID discovery, job object, Quit timing.

Questions:
1. Does each CoCreateInstanceEx(LOCAL_SERVER) start a *new* MSACCESS.EXE?
2. Does hWndAccessApp() work on a hidden, no-database instance, and does it resolve to MSACCESS.EXE?
3. Can a DCOM-launched Access be assigned to our job object? Does closing the job kill it?
4. Visible/UserControl defaults; bitness/runtime/version via SysCmd.
5. How long does Quit take to exit? Does holding a child COM reference keep the process alive after Quit?
"""

from __future__ import annotations

import time

import pywintypes
import win32api
import win32event
import win32job
from _harness import OwnedAccess

acSysCmdRuntime = 6
acSysCmdGetBitness = 724
acSysCmdGetFullVersion = 720


def make_kill_on_close_job() -> object:
    job = win32job.CreateJobObject(None, "")
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
    return job


def main() -> None:
    a = OwnedAccess()
    b = OwnedAccess()
    print(f"[1] A pid={a.pid} launch={a.launch_seconds:.2f}s  B pid={b.pid} launch={b.launch_seconds:.2f}s")
    print(f"    distinct instances: {a.pid != b.pid}")
    print(f"[2] A image={a.image}")
    print(f"[4] Visible={a.app.Visible} UserControl={a.app.UserControl} Version={a.app.Version} Build={a.app.Build}")
    for label, action in (("runtime", acSysCmdRuntime), ("bitness", acSysCmdGetBitness), ("fullversion", acSysCmdGetFullVersion)):
        try:
            print(f"    SysCmd({label}) = {a.app.SysCmd(action)!r}")
        except pywintypes.com_error as exc:
            print(f"    SysCmd({label}) failed: {exc}")
    print(f"    AutomationSecurity default = {a.app.AutomationSecurity}")

    # 3. job object assignment + kill-on-close
    job = make_kill_on_close_job()
    try:
        win32job.AssignProcessToJobObject(job, a.handle)
        print(f"[3] assigned A to job: in_job={win32job.IsProcessInJob(a.handle, job)}")
    except pywintypes.error as exc:
        print(f"[3] AssignProcessToJobObject failed: {exc}")

    # 5a. graceful quit timing for B while holding a child reference
    held = b.app.DBEngine  # child COM object kept alive across Quit
    started = time.perf_counter()
    b.app.Quit(2)
    b.app = None
    rc = win32event.WaitForSingleObject(b.handle, 15_000)
    print(f"[5] B exit after Quit while holding DBEngine ref: exited={rc == win32event.WAIT_OBJECT_0} after {time.perf_counter() - started:.2f}s")
    del held
    if rc != win32event.WAIT_OBJECT_0:
        rc = win32event.WaitForSingleObject(b.handle, 15_000)
        print(f"    after releasing ref: exited={rc == win32event.WAIT_OBJECT_0}")
        if rc != win32event.WAIT_OBJECT_0:
            b.quit()

    # 3b. closing the job handle must kill A (simulates Python dying)
    started = time.perf_counter()
    a.app = None  # drop proxy first so we measure the job effect, not a COM release
    win32api.CloseHandle(job) if isinstance(job, int) else job.Close()
    rc = win32event.WaitForSingleObject(a.handle, 15_000)
    print(f"[3] A killed by job close: {rc == win32event.WAIT_OBJECT_0} after {time.perf_counter() - started:.2f}s")
    if rc != win32event.WAIT_OBJECT_0:
        a.quit()

    # 5b. plain quit timing
    c = OwnedAccess()
    print(f"[5] C graceful quit took {c.quit():.2f}s")


if __name__ == "__main__":
    main()
