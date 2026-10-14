#!/usr/bin/env python3
"""Assemble the GitHub Pages site: docs/ plus the checker the web app runs in the browser.

    python scripts/build_site.py _site

The web app (docs/app/) loads the same scripts/ and rules/ files as the CLI, copied
unchanged into app/py/. Used by .github/workflows/pages.yml and for local testing
(python -m http.server -d _site).
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

COPIES = {
    "scripts/check_invoices.py": "app/py/scripts/check_invoices.py",
    "scripts/web_app.py": "app/py/scripts/web_app.py",
    "rules/itbis_rates.yaml": "app/py/rules/itbis_rates.yaml",
    "rules/ncf_types.yaml": "app/py/rules/ncf_types.yaml",
    "templates/plantilla_facturas.xlsx": "app/plantilla_facturas.xlsx",
    "examples/facturas_ejemplo.csv": "app/facturas_ejemplo.csv",
}


def build(out: Path) -> None:
    if out == ROOT or out in ROOT.parents or (out / ".git").exists():
        raise SystemExit(f"Refusing to overwrite {out}; pass a new output folder such as _site")
    if out.exists():
        shutil.rmtree(out)
    shutil.copytree(ROOT / "docs", out)
    for src, dst in COPIES.items():
        target = out / dst
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / src, target)
    (out / ".nojekyll").write_text("")


if __name__ == "__main__":
    build(Path(sys.argv[1] if len(sys.argv) > 1 else "_site").resolve())
