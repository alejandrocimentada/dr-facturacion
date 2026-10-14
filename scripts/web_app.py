"""Browser entry point for the web app (docs/app/), run inside Pyodide.

It calls the same functions as the CLI (scripts/check_invoices.py) and returns the results
as plain data for the page, plus the path of the Excel report it wrote. The file only
exists in the browser's in-memory file system; nothing is sent anywhere.

    result = check_file("/tmp/run1/facturas.xlsx", today="2026-10-14")
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import check_invoices as ci


def _display(value) -> str | float | None:
    """Cell value as shown in the Facturas sheet, made JSON-safe."""
    value = ci._cell_value(value, is_amount=False)
    if value is None:
        return None
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return float(value)
    return str(value)


def check_file(path: str, today: str | None = None) -> dict:
    """Check one CSV/XLSX file and write <name>_reporte.xlsx next to it."""
    input_path = Path(path)
    today_date = dt.date.fromisoformat(today) if today else dt.date.today()
    rules = ci.load_rules()
    try:
        df = ci.read_table(input_path)
        colmap = ci.detect_columns(list(df.columns))
    except ci.ColumnMappingError as exc:
        return {"ok": False, "error": "columns", "missing": exc.missing, "found": exc.found}
    except Exception as exc:  # unreadable file of any kind: tell the user, never crash the page
        return {"ok": False, "error": "read", "detail": str(exc)}
    if df.empty:
        return {"ok": False, "error": "empty"}

    result = ci.check_dataframe(df, colmap, rules, today_date)
    report = ci.write_report(result, df, input_path, ci.report_path_for(input_path), rules,
                             today_date, colmap)
    counts = result.counts()
    charged, expected = ci.itbis_totals(result)
    status = result.row_status()
    age = ci.rates_age_days(rules, today_date)
    return {
        "ok": True,
        "file": input_path.name,
        "today": today_date.strftime("%d/%m/%Y"),
        "invoices": result.invoice_count,
        "lines": len(result.lines),
        "lines_ok": sum(1 for l in result.lines if l.fila not in status),
        "counts": counts,
        "itbis": {"charged": round(charged, 2), "expected": round(expected, 2),
                  "diff": round(charged - expected, 2)},
        "rates": {"last_verified": str(rules["itbis"]["last_verified"]), "age_days": age,
                  "stale": age > ci.STALE_DAYS},
        "stdnum_overrides": result.stdnum_overrides,
        "issues": [{"fila": i.fila, "ncf": i.ncf, "severidad": i.severidad, "regla": i.regla,
                    "mensaje": i.mensaje, "filas": i.filas, "titulo": ci.RULE_HELP[i.regla][0],
                    "accion": ci.RULE_HELP[i.regla][1]}
                   for i in ci.sorted_issues(result.issues)],
        "actions": ci.top_actions(result),
        "columns": [str(c) for c in df.columns],
        "colmap": colmap,
        "rows": [{"fila": position + 2, "estado": status.get(position + 2, "OK"),
                  "values": [_display(v) for v in row.tolist()]}
                 for position, (_, row) in enumerate(df.iterrows())],
        "summary": ci.format_summary(result, input_path, None, rules, today_date),
        "report": str(report),
    }


def check_file_json(path: str, today: str | None = None) -> str:
    return json.dumps(check_file(path, today), ensure_ascii=False)

