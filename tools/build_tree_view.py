#!/usr/bin/env python3
"""リポジトリのファイルをツリー表示する HTML を生成する。

使い方:
    python tools/build_tree_view.py                 # tree_view.html を出力
    python tools/build_tree_view.py -o docs/out.html

Git の管理対象ファイル（`git ls-files`）だけを対象にし、テキストファイルは
中身ごと HTML に埋め込む。生成物は 1 枚の HTML なので、そのままブラウザで開ける。
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

# 中身を埋め込まない拡張子（バイナリ）
BINARY_EXT = {".ico", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".pdf", ".pptx", ".xlsx", ".zip"}
# 大きすぎるファイルは先頭だけ埋め込む
MAX_LINES = 400
MAX_BYTES = 400_000


def repo_root(start: Path) -> Path:
    out = subprocess.run(
        ["git", "rev-parse", "--show-toplevel"],
        cwd=start, capture_output=True, text=True, check=True,
    )
    return Path(out.stdout.strip())


def tracked_files(root: Path) -> list[str]:
    # -z を使うと日本語などの非 ASCII ファイル名がエスケープされずそのまま出る
    out = subprocess.run(
        ["git", "ls-files", "-z"], cwd=root, capture_output=True, check=True
    )
    names = out.stdout.decode("utf-8").split("\0")
    return sorted(name for name in names if name)


def current_branch(root: Path) -> str:
    out = subprocess.run(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"],
        cwd=root, capture_output=True, text=True,
    )
    return out.stdout.strip() or "HEAD"


def collect(root: Path, paths: list[str], skip: str | None = None) -> list[dict]:
    files = []
    for rel in paths:
        if rel == skip:  # 生成物そのものは埋め込まない
            continue
        abs_path = root / rel
        if not abs_path.is_file():
            continue
        size = abs_path.stat().st_size
        entry: dict = {"path": rel, "size": size, "lines": 0, "content": "", "binary": False}

        if abs_path.suffix.lower() in BINARY_EXT:
            entry["binary"] = True
            files.append(entry)
            continue

        try:
            text = abs_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            entry["binary"] = True
            files.append(entry)
            continue

        lines = text.splitlines()
        entry["lines"] = len(lines)
        if len(lines) > MAX_LINES or size > MAX_BYTES:
            text = "\n".join(lines[:MAX_LINES])
            entry["truncated"] = True
        entry["content"] = text
        files.append(entry)
    return files


def build(root: Path, template: Path, skip: str | None = None) -> str:
    data = {"branch": current_branch(root), "files": collect(root, tracked_files(root), skip)}
    payload = json.dumps(data, ensure_ascii=False)
    # <script> の中に置くため "<" を JSON のユニコードエスケープに逃がす。
    # これで "</script>" や "<!--" がタグとして解釈されることがなくなる。
    payload = payload.replace("<", "\\u003c")
    return template.read_text(encoding="utf-8").replace("__FILE_DATA__", payload)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="リポジトリのファイルツリー HTML を生成します。")
    parser.add_argument("-o", "--output", default="tree_view.html", help="出力先の HTML パス")
    args = parser.parse_args(argv)

    here = Path(__file__).resolve().parent
    root = repo_root(here)
    template = here / "tree_view_template.html"
    if not template.exists():
        print(f"テンプレートが見つかりません: {template}", file=sys.stderr)
        return 1

    output = Path(args.output)
    if not output.is_absolute():
        output = root / output
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        skip = output.resolve().relative_to(root).as_posix()
    except ValueError:
        skip = None
    output.write_text(build(root, template, skip), encoding="utf-8")
    print(f"生成しました: {output} ({output.stat().st_size:,} バイト)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
