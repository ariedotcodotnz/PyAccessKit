"""The ownership ledger's reaping rules (pure Python, fake process inspector)."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pyaccesskit import _ledger
from pyaccesskit._ledger import OwnedProcess


@dataclass
class FakeInspector:
    alive: set[tuple[int, float]] = field(default_factory=set)
    terminated: list[int] = field(default_factory=list)
    refuse: bool = False

    def matches(self, pid: int, creation_time: float) -> bool:
        return (pid, creation_time) in self.alive

    def terminate(self, pid: int, creation_time: float, image_name: str) -> bool:
        assert image_name == "MSACCESS.EXE"
        if self.refuse:
            return False
        self.alive.discard((pid, creation_time))
        self.terminated.append(pid)
        return True


def entry(pid: int, owner: int) -> OwnedProcess:
    return OwnedProcess(
        pid=pid,
        creation_time=1000.0 + pid,
        image=r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
        owner_pid=owner,
        owner_creation_time=500.0 + owner,
        database=None,
        started_at=0.0,
        version="test",
    )


def test_record_load_remove(tmp_path: Path) -> None:
    path = _ledger.record(entry(10, 1), tmp_path)
    assert [e for _p, e in _ledger.load_all(tmp_path)] == [entry(10, 1)]
    _ledger.remove(path)
    _ledger.remove(path)  # idempotent
    assert list(_ledger.load_all(tmp_path)) == []


def test_corrupt_entries_are_skipped(tmp_path: Path) -> None:
    (tmp_path / "junk.json").write_text("{not json", encoding="utf-8")
    _ledger.record(entry(10, 1), tmp_path)
    assert [e.pid for _p, e in _ledger.load_all(tmp_path)] == [10]


def test_reap_rules(tmp_path: Path) -> None:
    orphan, owned_by_live, exited = entry(10, 1), entry(20, 2), entry(30, 3)
    for item in (orphan, owned_by_live, exited):
        _ledger.record(item, tmp_path)
    inspector = FakeInspector(
        alive={
            (10, orphan.creation_time),
            (20, owned_by_live.creation_time),
            (2, owned_by_live.owner_creation_time),  # owner of 20 still runs
        }
    )
    dry = {
        r.entry.pid: r.action
        for r in _ledger.reap_orphans(inspector, dry_run=True, directory=tmp_path)
    }
    assert dry == {10: "would-terminate", 20: "owner-alive", 30: "already-exited"}
    assert inspector.terminated == []
    assert len(list(_ledger.load_all(tmp_path))) == 3, "a dry run changes nothing"

    done = {r.entry.pid: r.action for r in _ledger.reap_orphans(inspector, directory=tmp_path)}
    assert done == {10: "terminated", 20: "owner-alive", 30: "already-exited"}
    assert inspector.terminated == [10]
    assert [e.pid for _p, e in _ledger.load_all(tmp_path)] == [20]


def test_recycled_pid_is_never_terminated(tmp_path: Path) -> None:
    item = entry(10, 1)
    _ledger.record(item, tmp_path)
    inspector = FakeInspector(alive={(10, item.creation_time + 60)})  # same PID, different process
    results = _ledger.reap_orphans(inspector, directory=tmp_path)
    assert [r.action for r in results] == ["already-exited"]
    assert inspector.terminated == []


def test_failed_termination_keeps_the_entry(tmp_path: Path) -> None:
    item = entry(10, 1)
    _ledger.record(item, tmp_path)
    inspector = FakeInspector(alive={(10, item.creation_time)}, refuse=True)
    assert [r.action for r in _ledger.reap_orphans(inspector, directory=tmp_path)] == ["failed"]
    assert len(list(_ledger.load_all(tmp_path))) == 1


def test_ledger_dir_override(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.setenv(_ledger.LEDGER_ENV, str(tmp_path / "custom"))
    assert _ledger.ledger_dir() == tmp_path / "custom"
