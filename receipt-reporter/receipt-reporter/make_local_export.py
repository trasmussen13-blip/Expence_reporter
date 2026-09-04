from __future__ import annotations

import argparse
import zipfile
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parent
FILES = (
    "app.py",
    "requirements.txt",
    "pyproject.toml",
    "uv.lock",
    ".env.example",
    "LOCAL_SERVER.md",
    "run_local_server.sh",
    "make_local_export.py",
    "templates/index.html",
    "templates/login.html",
    "static/app.css",
    "static/favicon.svg",
    "receipt_data/base_template.xlsx",
)


def create_export(output: Path) -> Path:
    missing = [path for path in FILES if not (ROOT / path).is_file()]
    if missing:
        raise FileNotFoundError(f"Mangler filer til lokal eksport: {', '.join(missing)}")

    output.parent.mkdir(parents=True, exist_ok=True)
    package_name = "receipt-reporter"
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for relative_path in FILES:
            archive.write(ROOT / relative_path, f"{package_name}/{relative_path}")
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Opret ren ZIP til lokal server.")
    parser.add_argument(
        "--output",
        type=Path,
        default=ROOT / f"receipt-reporter-local-{date.today().isoformat()}.zip",
    )
    args = parser.parse_args()
    result = create_export(args.output)
    print(f"Oprettet: {result}")