"""Sync the package source into the agent skill scripts folder."""

from __future__ import annotations

import shutil
from pathlib import Path


def _project_root() -> Path:
    import mark2word

    return Path(mark2word.__file__).resolve().parents[2]


def sync_skill() -> Path:
    root = _project_root()
    src_pkg = root / "src" / "mark2word"
    scripts_dir = root / "skills" / "mark2word" / "scripts"
    dst_pkg = scripts_dir / "mark2word"
    if dst_pkg.exists():
        shutil.rmtree(dst_pkg)
    shutil.copytree(src_pkg, dst_pkg)
    launcher = scripts_dir / "mark2word.py"
    launcher.write_text(
        "import sys\n"
        "from pathlib import Path\n"
        "sys.path.insert(0, str(Path(__file__).resolve().parent))\n"
        "from mark2word.cli import main\n"
        "\n"
        "if __name__ == '__main__':\n"
        "    main()\n",
        encoding="utf-8",
    )
    return launcher


def main() -> None:
    dst = sync_skill()
    print(f"Synced skill launcher and package to {dst.parent}")


if __name__ == "__main__":
    main()
