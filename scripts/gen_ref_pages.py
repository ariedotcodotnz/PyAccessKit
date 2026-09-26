"""Generate one API reference page per public module (run by mkdocs-gen-files at build time).

Every module whose dotted path has no ``_``-prefixed part is public and gets a page rendering its
docstrings with mkdocstrings; the navigation is written to ``reference/api/SUMMARY.md`` for
mkdocs-literate-nav. The CLI is documented separately (docs/reference/cli.md).
"""

from __future__ import annotations

from pathlib import Path

import mkdocs_gen_files

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
SKIP = {("pyaccesskit", "cli"), ("pyaccesskit", "__main__")}

nav = mkdocs_gen_files.Nav()  # type: ignore[attr-defined]

for path in sorted((SRC / "pyaccesskit").rglob("*.py")):
    parts = tuple(path.relative_to(SRC).with_suffix("").parts)
    if parts[-1] == "__init__":
        parts = parts[:-1]
    if any(part.startswith("_") for part in parts) or parts[:2] in SKIP:
        continue
    is_package = path.name == "__init__.py"
    doc_path = Path(*parts, "index.md") if is_package else Path(*parts).with_suffix(".md")
    nav[parts] = doc_path.as_posix()
    identifier = ".".join(parts)
    with mkdocs_gen_files.open(Path("reference", "api", doc_path), "w") as page:
        if identifier == "pyaccesskit":
            # The top-level package re-exports everything; document it where it is defined.
            page.write(
                "# pyaccesskit\n\n"
                "::: pyaccesskit\n    options:\n      members: false\n      show_root_heading: false\n\n"
                "Everything listed in `pyaccesskit.__all__` is importable from the top-level package; "
                "the pages in this section document each object in the module that defines it.\n"
            )
        elif is_package:
            page.write(f"# {identifier}\n\n::: {identifier}\n    options:\n      members: false\n")
        else:
            page.write(f"# {identifier}\n\n::: {identifier}\n")
    mkdocs_gen_files.set_edit_path(Path("reference", "api", doc_path), path.relative_to(ROOT))

with mkdocs_gen_files.open("reference/api/SUMMARY.md", "w") as summary:
    summary.writelines(nav.build_literate_nav())
