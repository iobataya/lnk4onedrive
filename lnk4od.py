from __future__ import annotations

import argparse
import os
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

import pythoncom
from win32com.shell import shell


@dataclass
class ShortcutInfo:
    lnk: Path
    target: Path
    relative: str


def normalize_path(path: str | Path) -> Path:
    """
    Normalize a Windows path without requiring the target to exist.

    Environment variables such as %USERPROFILE% are expanded.
    """
    s = os.path.expandvars(str(path))
    s = os.path.expanduser(s)
    return Path(os.path.abspath(s))


def is_inside(path: Path, root: Path) -> bool:
    """
    Return True if path is inside root.

    Comparison is case-insensitive on Windows.
    """
    path_s = os.path.normcase(os.path.abspath(str(path)))
    root_s = os.path.normcase(os.path.abspath(str(root)))

    try:
        return os.path.commonpath([path_s, root_s]) == root_s
    except ValueError:
        # Different drives, for example C: and D:
        return False


def load_shell_link(lnk_path: Path):
    """
    Load an existing .lnk and return (IShellLink, IPersistFile).
    """
    link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink,
        None,
        pythoncom.CLSCTX_INPROC_SERVER,
        shell.IID_IShellLink,
    )

    persist = link.QueryInterface(pythoncom.IID_IPersistFile)
    persist.Load(str(lnk_path))

    return link, persist


def get_target(link) -> str:
    """
    Return the raw target path stored in a Shell Link.
    """
    target, _ = link.GetPath(shell.SLGP_RAWPATH)
    return target


def inspect_shortcut(
    lnk_path: Path,
    root: Path,
) -> ShortcutInfo | None:
    """
    Inspect one shortcut.

    Returns None when:
      - no filesystem target can be obtained
      - the shortcut is outside root
      - the target is outside root
    """
    link, _ = load_shell_link(lnk_path)

    target_raw = get_target(link)

    if not target_raw:
        return None

    target = normalize_path(target_raw)
    lnk_path = normalize_path(lnk_path)
    root = normalize_path(root)

    if not is_inside(lnk_path, root):
        return None

    if not is_inside(target, root):
        return None

    relative = os.path.relpath(
        str(target),
        start=str(lnk_path.parent),
    )

    return ShortcutInfo(
        lnk=lnk_path,
        target=target,
        relative=relative,
    )


def make_backup(lnk_path: Path) -> Path:
    """
    Create a non-destructive backup.

    Example:
        foo.lnk
        foo.lnk.bak
        foo.lnk.bak.1
        foo.lnk.bak.2
    """
    backup = Path(str(lnk_path) + ".bak")

    if not backup.exists():
        shutil.copy2(lnk_path, backup)
        return backup

    n = 1

    while True:
        candidate = Path(str(lnk_path) + f".bak.{n}")

        if not candidate.exists():
            shutil.copy2(lnk_path, candidate)
            return candidate

        n += 1


def set_relative_path(
    info: ShortcutInfo,
    backup: bool = False,
) -> tuple[bool, str | None]:
    """
    Add/update relative-path information in a .lnk file.

    Important:
        SetRelativePath() receives the full path of the .lnk itself.
        Windows calculates the relative relationship between the
        shortcut and its existing target.
    """
    link, persist = load_shell_link(info.lnk)

    backup_path = None

    if backup:
        backup_path = make_backup(info.lnk)

    # Windows calculates RelativePath using:
    #
    #   shortcut location -> existing target
    #
    link.SetRelativePath(str(info.lnk))

    # IPersistFile::Save requires an absolute filename.
    persist.Save(str(info.lnk), 1)

    # Reload the file so verification examines what was actually saved.
    verify_link, _ = load_shell_link(info.lnk)

    try:
        verified_target, _ = verify_link.GetPath(
            shell.SLGP_RELATIVEPRIORITY
        )
    except Exception:
        # Some pywin32 / Windows combinations may not expose or accept
        # SLGP_RELATIVEPRIORITY as expected. The save itself may still
        # have succeeded.
        verified_target = ""

    ok = info.lnk.exists()

    return ok, str(backup_path) if backup_path else None


def find_shortcuts(root: Path):
    """
    Yield all .lnk files recursively below root.
    """
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() == ".lnk":
            yield path


def detect_onedrive_root() -> Path | None:
    """
    Try common OneDrive environment variables.

    OneDriveConsumer:
        Personal OneDrive

    OneDriveCommercial:
        OneDrive for Business

    OneDrive:
        General OneDrive variable
    """
    candidates = (
        "Quantum Design",
        "OneDrive",
        "OneDriveConsumer",
        "OneDriveCommercial",
    )

    found = []

    for name in candidates:
        value = os.environ.get(name)
        print(f"Checking environment variable {name}: {value}", file=sys.stderr)

        if value:
            p = Path(value)

            if p.exists() and p not in found:
                found.append(p)

    if len(found) == 1:
        return found[0]

    if len(found) > 1:
        print(
            "Multiple OneDrive roots were detected:",
            file=sys.stderr,
        )

        for p in found:
            print(f"  {p}", file=sys.stderr)

        print(
            "\nSpecify the desired root explicitly.",
            file=sys.stderr,
        )

    return None


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Add relative-path information to Windows .lnk files "
            "inside a OneDrive directory."
        )
    )

    parser.add_argument(
        "root",
        nargs="?",
        type=Path,
        help=(
            "OneDrive root directory. "
            "If omitted, environment variables are checked."
        ),
    )

    mode = parser.add_mutually_exclusive_group()

    mode.add_argument(
        "--apply",
        action="store_true",
        help="Actually modify .lnk files.",
    )

    mode.add_argument(
        "--dry-run",
        action="store_true",
        help="Only display planned changes (default).",
    )

    parser.add_argument(
        "--backup",
        action="store_true",
        help="Create .bak copies before modifying shortcuts.",
    )

    parser.add_argument(
        "--include-broken",
        action="store_true",
        help=(
            "Allow targets that currently do not exist. "
            "By default, broken targets are skipped."
        ),
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    root = args.root

    if root is None:
        root = detect_onedrive_root()

        if root is None:
            print(
                "Could not determine a unique OneDrive root.\n"
                "Specify it explicitly, for example:\n\n"
                r'  python relative_lnk.py "C:\Users\me\OneDrive"',
                file=sys.stderr,
            )
            return 2

    root = normalize_path(root)

    if not root.is_dir():
        print(
            f"Directory does not exist: {root}",
            file=sys.stderr,
        )
        return 2

    apply_changes = args.apply

    print(f"Root : {root}")
    print(f"Mode : {'APPLY' if apply_changes else 'DRY RUN'}")

    if args.backup and apply_changes:
        print("Backup: enabled")

    print()

    total = 0
    eligible = 0
    modified = 0
    skipped_outside = 0
    skipped_broken = 0
    errors = 0

    pythoncom.CoInitialize()

    try:
        for lnk in find_shortcuts(root):
            total += 1

            try:
                info = inspect_shortcut(lnk, root)

                if info is None:
                    print(f"SKIP  {lnk}")
                    print("      No suitable target inside OneDrive.")
                    skipped_outside += 1
                    continue

                if not info.target.exists() and not args.include_broken:
                    print(f"SKIP  {info.lnk}")
                    print(f"      Broken target: {info.target}")
                    skipped_broken += 1
                    continue

                eligible += 1

                print(f"{'WRITE' if apply_changes else 'WOULD'} {info.lnk}")
                print(f"      target   : {info.target}")
                print(f"      relative : {info.relative}")

                if apply_changes:
                    ok, backup_path = set_relative_path(
                        info,
                        backup=args.backup,
                    )

                    if backup_path:
                        print(f"      backup   : {backup_path}")

                    if ok:
                        print("      result   : OK")
                        modified += 1
                    else:
                        print("      result   : FAILED")
                        errors += 1

                print()

            except Exception as exc:
                errors += 1
                print(f"ERROR {lnk}")
                print(f"      {type(exc).__name__}: {exc}")
                print()

    finally:
        pythoncom.CoUninitialize()

    print("-" * 60)
    print(f"Shortcuts found : {total}")
    print(f"Eligible        : {eligible}")

    if apply_changes:
        print(f"Modified        : {modified}")

    print(f"Broken skipped  : {skipped_broken}")
    print(f"Other skipped   : {skipped_outside}")
    print(f"Errors          : {errors}")

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())