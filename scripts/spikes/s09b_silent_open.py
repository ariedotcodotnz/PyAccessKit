"""S9b — does OpenCurrentDatabase fail *silently* (no exception) for locked / non-database files?"""

from __future__ import annotations

import time

import pywintypes
from _harness import DialogWatcher, describe_com_error, fresh_dir, owned_access


def state(app) -> str:
    parts = []
    try:
        parts.append(f"FullName={app.CurrentProject.FullName!r}")
    except pywintypes.com_error as exc:
        parts.append(f"FullName error {describe_com_error(exc)}")
    try:
        db = app.CurrentDb()
        parts.append(f"CurrentDb={'None' if db is None else db.Name!r}")
    except pywintypes.com_error as exc:
        parts.append(f"CurrentDb error {describe_com_error(exc)}")
    return "; ".join(parts)


def main() -> None:
    work = fresh_dir("s09b")
    path = work / "locked.accdb"
    bogus = work / "text.accdb"
    bogus.write_text("not a database", encoding="ascii")
    with owned_access() as holder:
        holder.app.NewCurrentDatabase(str(path), 12)
        holder.app.CloseCurrentDatabase()
        holder.app.OpenCurrentDatabase(str(path), True)  # exclusive
        with owned_access() as other, DialogWatcher(other.pid) as watch:
            for label, target, exclusive in (
                ("locked by other instance (shared)", path, False),
                ("locked by other instance (exclusive)", path, True),
                ("text file named .accdb", bogus, False),
            ):
                t0 = time.perf_counter()
                try:
                    other.app.OpenCurrentDatabase(str(target), exclusive)
                    print(f"{label}: no exception after {time.perf_counter() - t0:.2f}s -> {state(other.app)}")
                except pywintypes.com_error as exc:
                    print(f"{label}: raised {describe_com_error(exc)}")
                time.sleep(1)
                try:
                    other.app.CloseCurrentDatabase()
                except pywintypes.com_error as exc:
                    print(f"   CloseCurrentDatabase: {describe_com_error(exc)}")
            print(f"dialogs: {watch.events}")
        holder.app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
