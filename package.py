"""Assemble a clean Windows distribution folder + zip.

Refreshes the front-end workbook, copies only the runtime files (no venv,
caches, dbs or sample outputs), normalises .bat/.txt/.bas to CRLF so they
behave on Windows, and zips the result. Run: python package.py
"""
import shutil
import zipfile
from pathlib import Path

import build_frontend

BASE = Path(__file__).resolve().parent
DIST = BASE / "dist"
PKG = DIST / "PortfolioTool"

# runtime files to ship (everything needed to install, verify, and use)
FILES = [
    "app.py", "config.py", "db.py", "ingest.py", "compute.py", "charts.py",
    "build_frontend.py", "make_samples.py", "test_pipeline.py",
    "FundTool.bas", "requirements.txt", "install.bat", "run_headless.bat",
    "START_HERE.txt", "SETUP.md", "README.md", "PortfolioTool.xlsx",
]
DIRS = ["incoming_samples"]                 # sample fund files (3 types)
CRLF_EXT = {".bat", ".txt", ".bas"}         # Windows-sensitive text files


def _copy_text_crlf(src: Path, dst: Path):
    text = src.read_text(encoding="utf-8")
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    dst.write_bytes(text.encode("utf-8"))


def main():
    print("Refreshing PortfolioTool.xlsx ...")
    build_frontend.main()

    if DIST.exists():
        shutil.rmtree(DIST)
    PKG.mkdir(parents=True)

    for name in FILES:
        src = BASE / name
        if not src.exists():
            raise FileNotFoundError(src)
        dst = PKG / name
        if src.suffix in CRLF_EXT:
            _copy_text_crlf(src, dst)
        else:
            shutil.copy2(src, dst)

    for d in DIRS:
        for src in (BASE / d).rglob("*"):
            if src.is_file() and "__pycache__" not in src.parts:
                rel = src.relative_to(BASE)
                dst = PKG / rel
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)

    zip_path = DIST / "PortfolioTool.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(PKG.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(DIST))

    n = sum(1 for _ in PKG.rglob("*") if _.is_file())
    size_kb = zip_path.stat().st_size / 1024
    print(f"\nPackaged {n} files -> {PKG}")
    print(f"Zip: {zip_path}  ({size_kb:.0f} KB)")


if __name__ == "__main__":
    main()
