#!/usr/bin/env python3
"""Create a reproducible Linux sidecar tarball while preserving executable modes."""

from __future__ import annotations

import gzip
import sys
import tarfile
from pathlib import Path


def normalize(info: tarfile.TarInfo) -> tarfile.TarInfo:
    info.uid = 0
    info.gid = 0
    info.uname = "root"
    info.gname = "root"
    info.mtime = 0
    normalized = info.name.replace("\\", "/")
    if info.isdir() or normalized.endswith("/node/bin/node"):
        info.mode = 0o755
    elif info.isfile():
        info.mode = 0o644
    return info


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("usage: package_linux_sidecar.py PACKAGE_DIR OUTPUT.tar.gz")
    source = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]).resolve()
    if not source.is_dir() or not (source / "node" / "bin" / "node").is_file():
        raise SystemExit(f"invalid Linux sidecar tree: {source}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("wb") as raw:
        with gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=0) as compressed:
            with tarfile.open(fileobj=compressed, mode="w", format=tarfile.PAX_FORMAT) as archive:
                archive.add(source, arcname="vector-sidecar", filter=normalize)


if __name__ == "__main__":
    main()
