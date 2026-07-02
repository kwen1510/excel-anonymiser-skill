#!/usr/bin/env python3
"""Install the Excel privacy handler template into a target project."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = SKILL_DIR / "assets" / "template"


def copy_template(target: Path, force: bool) -> None:
    if not TEMPLATE_DIR.exists():
        raise SystemExit(f"Template directory missing: {TEMPLATE_DIR}")
    target.mkdir(parents=True, exist_ok=True)
    for item in TEMPLATE_DIR.iterdir():
        destination = target / item.name
        if destination.exists() and not force:
            raise SystemExit(f"Refusing to overwrite existing path without --force: {destination}")
        if destination.exists():
            if destination.is_dir():
                shutil.rmtree(destination)
            else:
                destination.unlink()
        if item.is_dir():
            shutil.copytree(item, destination)
        else:
            shutil.copy2(item, destination)


def main() -> int:
    parser = argparse.ArgumentParser(description="Install the Excel privacy handler template")
    parser.add_argument("--target", required=True, help="Project directory to receive the scaffold")
    parser.add_argument("--force", action="store_true", help="Replace existing scaffold files")
    args = parser.parse_args()
    copy_template(Path(args.target).expanduser().resolve(), args.force)
    print(f"Installed Excel privacy handler template into {Path(args.target).expanduser().resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
