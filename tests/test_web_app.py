"""The web app's Python entry point must give exactly the same results as the CLI."""

import json
import shutil
from pathlib import Path

import check_invoices as ci
import pandas as pd
import web_app

ROOT = Path(__file__).resolve().parent.parent
EXAMPLE = ROOT / "examples" / "facturas_ejemplo.csv"
TODAY = "2026-10-07"


def run_both(tmp_path, src=EXAMPLE):
    cli_dir, web_dir = tmp_path / "cli", tmp_path / "web"
    cli_dir.mkdir()
    web_dir.mkdir()
    cli_input = shutil.copy(src, cli_dir / src.name)
    web_input = shutil.copy(src, web_dir / src.name)
    assert ci.main([str(cli_input), "--today", TODAY]) == 0
    cli_report = ci.report_path_for(Path(cli_input))
    return cli_report, web_app.check_file(str(web_input), today=TODAY)


def test_same_problems_as_cli(tmp_path):
    cli_report, web = run_both(tmp_path)
    problems = pd.read_excel(cli_report, sheet_name="Problemas", keep_default_na=False)
    cli_rows = [tuple(r) for r in problems[["fila", "ncf", "severidad", "regla", "mensaje"]].itertuples(index=False)]
    web_rows = [(i["fila"], i["ncf"], i["severidad"], i["regla"], i["mensaje"]) for i in web["issues"]]
    assert web_rows == cli_rows
    assert set(i["regla"] for i in web["issues"]) == set(ci.RULES)


def test_same_row_status_as_cli(tmp_path):
    cli_report, web = run_both(tmp_path)
    facturas = pd.read_excel(cli_report, sheet_name="Facturas")
    assert [r["estado"] for r in web["rows"]] == list(facturas["estado"])
    assert [r["fila"] for r in web["rows"]] == list(facturas["fila"])
    assert web["columns"] == list(facturas.columns[1:-1])


def test_same_summary_numbers_as_cli(tmp_path, capsys):
    _, web = run_both(tmp_path)
    cli_out = capsys.readouterr().out
    assert web["counts"] == {"ERROR": 15, "ADVERTENCIA": 4, "INFO": 4}
    assert f"Facturas revisadas: {web['invoices']} ({web['lines']} líneas)" in cli_out
    assert "ERROR: 15 | ADVERTENCIA: 4 | INFO: 4" in cli_out
    assert web["itbis"] == {"charged": 44098.8, "expected": 44258.8, "diff": -160.0}
    # the terminal summary is identical apart from the report path line
    cli_lines = [l for l in cli_out.splitlines() if not l.startswith("Reporte:")]
    assert web["summary"].splitlines() == cli_lines


def test_same_excel_report_as_cli(tmp_path):
    cli_report, web = run_both(tmp_path)
    web_report = Path(web["report"])
    assert web_report.name == "facturas_ejemplo_reporte.xlsx"
    cli_sheets = pd.read_excel(cli_report, sheet_name=None, header=None)
    web_sheets = pd.read_excel(web_report, sheet_name=None, header=None)
    assert list(web_sheets) == list(cli_sheets) == ["Resumen", "Problemas", "Facturas"]
    for name in cli_sheets:
        pd.testing.assert_frame_equal(web_sheets[name], cli_sheets[name])


def test_xlsx_input_matches_cli(tmp_path):
    xlsx = tmp_path / "facturas.xlsx"
    pd.read_csv(EXAMPLE, dtype=str, keep_default_na=False).to_excel(xlsx, index=False)
    cli_report, web = run_both(tmp_path, xlsx)
    problems = pd.read_excel(cli_report, sheet_name="Problemas")
    assert [i["regla"] for i in web["issues"]] == list(problems["regla"])


def test_result_is_json_serializable(tmp_path):
    src = shutil.copy(EXAMPLE, tmp_path / "f.csv")
    data = json.loads(web_app.check_file_json(str(src), today=TODAY))
    assert data["ok"] and data["file"] == "f.csv" and len(data["rows"]) == 28
    assert len(data["actions"]) == 3 and data["actions"][0]["severidad"] == "ERROR"


def test_unknown_columns_are_reported_not_raised(tmp_path):
    src = tmp_path / "raro.csv"
    src.write_text("a;b;c\n1;2;3\n", encoding="utf-8")
    result = web_app.check_file(str(src), today=TODAY)
    assert result["ok"] is False and result["error"] == "columns"
    assert "ncf" in result["missing"] and result["found"] == ["a", "b", "c"]
    assert not (tmp_path / "raro_reporte.xlsx").exists()


def test_unreadable_file_is_reported(tmp_path):
    src = tmp_path / "roto.xlsx"
    src.write_bytes(b"not a real workbook")
    result = web_app.check_file(str(src), today=TODAY)
    assert result == {"ok": False, "error": "read", "detail": result["detail"]}
