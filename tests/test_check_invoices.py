import datetime as dt
import hashlib
import shutil
from pathlib import Path

import check_invoices as ci
import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parent.parent

TODAY = dt.date(2026, 10, 7)
RULES = ci.load_rules()

# Documented valid examples from python-stdnum (stdnum.do.rnc / stdnum.do.cedula).
RNC_OK = "1-01-85004-3"
RNC_BAD = "1-01-85004-2"          # same number, wrong check digit
RNC_OK_2 = "131246796"
CEDULA_OK = "00113918205"
CEDULA_BAD = "00113918204"        # wrong check digit
BUYER_RNC = "130123454"           # check digit computed with stdnum.do.rnc.calc_check_digit


def line(**overrides):
    base = {
        "fecha": "15/09/2026", "rnc_emisor": RNC_OK, "rnc_comprador": BUYER_RNC,
        "ncf": "B0100000001", "descripcion": "Servicio", "cantidad": "1",
        "precio_unitario": "1000.00", "tasa_itbis": "18", "itbis": "180.00", "total": "1180.00",
        "vencimiento_secuencia": "31/12/2026", "ncf_modificado": "",
    }
    base.update(overrides)
    return base


def run(*rows, drop=()):
    df = pd.DataFrame(list(rows)).drop(columns=list(drop))
    colmap = ci.detect_columns(list(df.columns))
    return ci.check_dataframe(df, colmap, RULES, TODAY)


def codes(result):
    return [i.regla for i in result.issues]


def test_clean_invoice_has_no_issues():
    assert codes(run(line())) == []


# ---------------------------------------------------------------- 1. RNC / cédula

@pytest.mark.parametrize("value,kind", [
    (RNC_OK, "rnc"), ("101850043", "rnc"), (" 1 01 85004 3 ", "rnc"), (RNC_OK_2, "rnc"),
    (CEDULA_OK, "cedula"), ("001-1391820-5", "cedula"), ("224-0002211-1", "cedula"),
    (101850043, "rnc"), (101850043.0, "rnc"), ("", "vacio"), (None, "vacio"),
    (RNC_BAD, "invalido"), (CEDULA_BAD, "invalido"), ("1018A0043", "invalido"),
    ("12345", "invalido"), ("1234567890123", "invalido"),
    ("113918205", "cedula_ceros"), (113918205, "cedula_ceros"),
])
def test_classify_id(value, kind):
    assert ci.classify_id(value)[0] == kind


def test_invalid_emitter_rnc_is_error():
    result = run(line(rnc_emisor=RNC_BAD))
    assert codes(result) == ["R1-RNC-EMISOR"]
    assert result.issues[0].severidad == ci.ERROR


def test_missing_emitter_is_error():
    assert "R1-RNC-EMISOR" in codes(run(line(rnc_emisor="")))


def test_invalid_buyer_cedula_is_error():
    assert "R1-RNC-COMPRADOR" in codes(run(line(ncf="B0200000001", rnc_comprador=CEDULA_BAD)))


def test_wrong_length_buyer_is_error():
    result = run(line(ncf="B0200000001", rnc_comprador="1234567"))
    assert codes(result) == ["R1-RNC-COMPRADOR"]
    assert "7 dígitos" in result.issues[0].mensaje


def test_cedula_missing_leading_zeros_is_info_and_accepted():
    result = run(line(rnc_comprador="113918205"))
    assert codes(result) == ["R1-CEDULA-CEROS"]
    assert result.issues[0].severidad == ci.INFO


# ---------------------------------------------------------------- 2. NCF structure

@pytest.mark.parametrize("ncf", ["B0100000001", "b01-0000-0001", "E310000000001", "B1700000099",
                                 "E470000000001", "B1200000001", "E410000000001"])
def test_valid_ncf(ncf):
    info, code, _ = ci.parse_ncf(ncf, RULES["ncf"])
    assert code is None and info.text == ncf.upper().replace("-", "")


@pytest.mark.parametrize("ncf,code", [
    ("", "R2-NCF-ESTRUCTURA"),
    ("B010000001", "R2-NCF-ESTRUCTURA"),       # 10 chars
    ("B01000000001", "R2-NCF-ESTRUCTURA"),     # 12 chars
    ("E3100000001", "R2-NCF-ESTRUCTURA"),      # e-NCF with 11 chars
    ("B01000000A1", "R2-NCF-ESTRUCTURA"),      # letter in sequence
    ("Z0100000001", "R2-NCF-ESTRUCTURA"),      # unknown series
    ("A020010210100000005", "R2-NCF-ESTRUCTURA"),  # pre-2018 format
    ("B0900000001", "R2-NCF-TIPO"),
    ("B3100000001", "R2-NCF-TIPO"),            # electronic type on a B series
    ("E010000000001", "R2-NCF-TIPO"),          # paper type on an E series
    ("E420000000001", "R2-NCF-TIPO"),
])
def test_invalid_ncf(ncf, code):
    result = run(line(ncf=ncf))
    assert code in codes(result)
    assert ci.RULES[code][0] == ci.ERROR


def test_yaml_is_trusted_over_stdnum(monkeypatch):
    monkeypatch.setattr(ci.std_ncf, "is_valid", lambda number: False)
    result = run(line(ncf="E310000000001"))
    assert codes(result) == []
    assert result.stdnum_overrides == ["E31"]


def test_stdnum_and_yaml_agree_on_all_types():
    for series, types in RULES["ncf"]["types"].items():
        width = RULES["ncf"]["series"][series]["sequence_digits"]
        for tipo in types:
            assert ci.std_ncf.is_valid(f"{series}{tipo}{1:0{width}d}"), series + tipo


# ---------------------------------------------------------------- 3. NCF type vs buyer

@pytest.mark.parametrize("ncf", ["B0100000001", "E310000000001"])
def test_credito_fiscal_without_buyer_is_error(ncf):
    result = run(line(ncf=ncf, rnc_comprador=""))
    assert codes(result) == ["R3-CREDITO-SIN-COMPRADOR"]


def test_credito_fiscal_with_invalid_buyer_is_error():
    assert set(codes(run(line(rnc_comprador=RNC_BAD)))) == {"R1-RNC-COMPRADOR", "R3-CREDITO-SIN-COMPRADOR"}


def test_credito_fiscal_with_cedula_is_ok():
    assert codes(run(line(ncf="E310000000001", rnc_comprador=CEDULA_OK))) == []


@pytest.mark.parametrize("ncf", ["B0200000001", "E320000000001"])
def test_consumo_with_buyer_rnc_is_warning(ncf):
    result = run(line(ncf=ncf, rnc_comprador=BUYER_RNC))
    assert codes(result) == ["R3-CONSUMO-CON-RNC"]
    assert result.issues[0].severidad == ci.ADVERTENCIA


@pytest.mark.parametrize("buyer", ["", CEDULA_OK])
def test_consumo_without_rnc_is_ok(buyer):
    assert codes(run(line(ncf="B0200000001", rnc_comprador=buyer))) == []


@pytest.mark.parametrize("ncf", ["B1400000001", "E440000000001", "B1600000001", "E460000000001"])
def test_special_regime_or_export_with_itbis_is_warning(ncf):
    assert codes(run(line(ncf=ncf))) == ["R3-EXENTO-CON-ITBIS"]


@pytest.mark.parametrize("rate", ["0", "exento"])
def test_special_regime_without_itbis_is_ok(rate):
    assert codes(run(line(ncf="B1600000001", tasa_itbis=rate, itbis="0", total="1000"))) == []


@pytest.mark.parametrize("ncf", ["B0300000001", "B0400000001", "E330000000001", "E340000000001"])
def test_note_without_modified_ncf_is_warning(ncf):
    assert codes(run(line(ncf=ncf))) == ["R3-NOTA-SIN-REFERENCIA"]


def test_note_with_modified_ncf_is_ok():
    assert codes(run(line(ncf="B0400000001", ncf_modificado="B0100000001"))) == []


def test_note_when_column_missing_is_warning():
    assert codes(run(line(ncf="B0400000001"), drop=["ncf_modificado"])) == ["R3-NOTA-SIN-REFERENCIA"]


# ---------------------------------------------------------------- 4. Sequences

def test_multi_line_invoice_is_not_duplicate():
    result = run(line(descripcion="A"), line(descripcion="B"))
    assert codes(result) == [] and result.invoice_count == 1


def test_duplicate_ncf_is_error():
    result = run(line(), line(fecha="20/09/2026"))
    assert codes(result) == ["R4-NCF-DUPLICADO"]
    assert result.issues[0].fila == 3 and result.invoice_count == 2


def test_duplicate_ncf_with_other_buyer_is_error():
    assert codes(run(line(), line(rnc_comprador=CEDULA_OK))) == ["R4-NCF-DUPLICADO"]


def test_same_ncf_different_emitters_is_ok():
    assert codes(run(line(), line(rnc_emisor=RNC_OK_2))) == []


def test_identical_lines_are_warning():
    result = run(line(), line())
    assert "R4-LINEA-REPETIDA" in codes(result)


def test_sequence_gap_is_info():
    result = run(line(ncf="B0100000001"), line(ncf="B0100000004"))
    assert codes(result) == ["R4-SECUENCIA-SALTO"]
    issue = result.issues[0]
    assert issue.severidad == ci.INFO and "B0100000002 a B0100000003" in issue.mensaje


def test_gaps_are_per_emitter_and_type():
    result = run(line(ncf="B0100000001"), line(ncf="B0100000002"),
                 line(ncf="B0200000007", rnc_comprador=""),
                 line(ncf="B0100000005", rnc_emisor=RNC_OK_2))
    assert codes(result) == []


def test_expired_sequence_is_error():
    result = run(line(fecha="01/10/2026", vencimiento_secuencia="30/09/2026"))
    assert codes(result) == ["R4-SECUENCIA-VENCIDA"]


def test_sequence_not_expired_on_last_day():
    assert codes(run(line(fecha="30/09/2026", vencimiento_secuencia="30/09/2026"))) == []


def test_expiry_column_optional():
    assert codes(run(line(), drop=["vencimiento_secuencia"])) == []


# ---------------------------------------------------------------- 5. ITBIS math

@pytest.mark.parametrize("value,key", [
    ("18", "general"), ("18%", "general"), ("0.18", "general"), (18, "general"), (18.0, "general"),
    ("16", "reducida"), ("0", "exportacion"), ("0%", "exportacion"),
    ("exento", "exento"), ("Exento", "exento"), ("EXENTA", "exento"), ("E", "exento"),
])
def test_parse_rate(value, key):
    assert ci.parse_rate(value, RULES["itbis"]["rates"])[0] == key


@pytest.mark.parametrize("rate", ["15", "8%", "", "general", "0.19"])
def test_invalid_rate_is_error(rate):
    assert "R5-TASA-INVALIDA" in codes(run(line(tasa_itbis=rate)))


def test_line_itbis_mismatch_is_error_with_expected_value():
    result = run(line(itbis="170.00", total="1170.00"))
    assert codes(result) == ["R5-ITBIS-LINEA"]
    assert "RD$170.00" in result.issues[0].mensaje and "RD$180.00" in result.issues[0].mensaje


def test_line_itbis_within_one_cent_is_ok():
    assert codes(run(line(itbis="180.01", total="1180.01"))) == []


def test_line_itbis_two_cents_off_is_error():
    assert codes(run(line(itbis="180.02", total="1180.02"))) == ["R5-ITBIS-LINEA"]


def test_line_itbis_uses_quantity():
    assert codes(run(line(cantidad="3", precio_unitario="333.33", itbis="179.99", total="1179.98"))) == []


def test_exempt_line_with_itbis_is_error():
    result = run(line(tasa_itbis="exento"))
    assert codes(result) == ["R5-ITBIS-LINEA"]
    assert "exenta" in result.issues[0].mensaje


def test_invoice_total_mismatch_is_error():
    assert codes(run(line(total="1200.00"))) == ["R5-TOTAL-FACTURA"]


def test_invoice_total_within_five_cents_is_ok():
    assert codes(run(line(total="1180.05"))) == []


def test_invoice_total_is_summed_over_lines():
    rows = [line(descripcion="A"), line(descripcion="B", total="1180.10")]
    assert codes(run(*rows)) == ["R5-TOTAL-FACTURA"]


def test_invoice_total_repeated_on_every_line_is_accepted():
    rows = [line(descripcion="A", total="2360.00"), line(descripcion="B", total="2360.00")]
    assert codes(run(*rows)) == []


def test_reduced_rate_is_info():
    result = run(line(tasa_itbis="16", itbis="160.00", total="1160.00"))
    assert codes(result) == ["R5-TASA-REDUCIDA"]
    assert result.issues[0].severidad == ci.INFO


@pytest.mark.parametrize("field", ["cantidad", "precio_unitario", "total"])
def test_unreadable_or_missing_number_is_error(field):
    assert "R5-NUMERO-INVALIDO" in codes(run(line(**{field: "abc"})))
    assert "R5-NUMERO-INVALIDO" in codes(run(line(**{field: ""})))


def test_blank_itbis_means_zero():
    assert codes(run(line(tasa_itbis="0", itbis="", total="1000"))) == []


@pytest.mark.parametrize("text,value", [
    ("1180", 1180.0), ("1,180.00", 1180.0), ("RD$ 1,180.50", 1180.5), ("1.180,50", 1180.5),
    ("1180,5", 1180.5), ("(100.00)", -100.0), (1180, 1180.0), ("", None),
])
def test_parse_number(text, value):
    assert ci.parse_number(text) == value


# ---------------------------------------------------------------- 6. Dates

@pytest.mark.parametrize("value", ["31/02/2026", "2026-13-01", "ayer", "15-Sep"])
def test_unparseable_date_is_error(value):
    assert codes(run(line(fecha=value))) == ["R6-FECHA-INVALIDA"]


def test_missing_date_is_error():
    assert "R6-FECHA-INVALIDA" in codes(run(line(fecha="")))


def test_future_date_is_error():
    assert codes(run(line(fecha="08/10/2026"))) == ["R6-FECHA-FUTURA"]


def test_today_is_not_future():
    assert codes(run(line(fecha="07/10/2026"))) == []


@pytest.mark.parametrize("value", ["15/09/2026", "2026-09-15", "2026-09-15 00:00:00", "15-09-2026",
                                   dt.datetime(2026, 9, 15), pd.Timestamp("2026-09-15"), 46280])
def test_parse_date_formats(value):
    assert ci.parse_date(value) == dt.date(2026, 9, 15)


# ---------------------------------------------------------------- columns

def test_detect_spanish_headers_with_accents():
    headers = ["Fecha", "RNC Emisor", "RNC/Cédula Comprador", "NCF", "Descripción", "Cantidad",
               "Precio Unitario", "Tasa ITBIS", "ITBIS", "Total", "Vencimiento Secuencia", "NCF Modificado"]
    mapping = ci.detect_columns(headers)
    assert mapping["rnc_comprador"] == "RNC/Cédula Comprador"
    assert mapping["descripcion"] == "Descripción" and mapping["ncf_modificado"] == "NCF Modificado"


def test_detect_english_headers():
    headers = ["Date", "Issuer RNC", "Buyer RNC", "e-NCF", "Description", "Qty", "Unit Price",
               "ITBIS %", "Tax Amount", "Line Total"]
    mapping = ci.detect_columns(headers)
    assert mapping["tasa_itbis"] == "ITBIS %" and mapping["itbis"] == "Tax Amount"
    assert mapping["ncf"] == "e-NCF" and mapping["total"] == "Line Total"


def test_missing_columns_raise_and_manual_map_fixes_them():
    headers = ["Fecha Doc", "RNC Emisor", "Cliente", "No. Comprobante", "Cantidad", "Precio",
               "Tasa", "ITBIS", "Total"]
    with pytest.raises(ci.ColumnMappingError) as exc:
        ci.detect_columns(headers)
    assert set(exc.value.missing) == {"fecha", "rnc_comprador"}
    mapping = ci.detect_columns(headers, {"fecha": "Fecha Doc", "rnc_comprador": "cliente"})
    assert mapping["fecha"] == "Fecha Doc" and mapping["rnc_comprador"] == "Cliente"


def test_manual_map_to_unknown_header_fails():
    with pytest.raises(ValueError):
        ci.detect_columns(["a"], {"ncf": "Nope"})


def test_template_headers_are_detected():
    df = pd.read_excel(ROOT / "templates" / "plantilla_facturas.xlsx", dtype=object)
    mapping = ci.detect_columns(list(df.columns))
    assert set(ci.REQUIRED_COLUMNS + ci.OPTIONAL_COLUMNS) == set(mapping)


# ---------------------------------------------------------------- rates freshness

def test_rates_file_values():
    rates = RULES["itbis"]["rates"]
    assert rates == {"general": 0.18, "reducida": 0.16, "exportacion": 0.0, "exento": None}
    assert len(RULES["itbis"]["sources"]) >= 2


def test_stale_rates_warning():
    result = run(line())
    fresh = ci.format_summary(result, Path("x.csv"), None, RULES, dt.date(2026, 10, 7))
    stale = ci.format_summary(result, Path("x.csv"), None, RULES, dt.date(2027, 1, 6))
    assert "AVISO" not in fresh
    assert "AVISO" in stale and "91 días" in stale


# ---------------------------------------------------------------- end to end

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_example_file_end_to_end(tmp_path, capsys):
    src = tmp_path / "facturas_ejemplo.csv"
    shutil.copy(ROOT / "examples" / "facturas_ejemplo.csv", src)
    before = sha(src)
    assert ci.main([str(src), "--today", "2026-10-07"]) == 0
    assert sha(src) == before  # input never modified

    report = tmp_path / "facturas_ejemplo_reporte.xlsx"
    sheets = pd.read_excel(report, sheet_name=None)
    assert list(sheets) == ["Resumen", "Problemas", "Facturas"]
    assert list(sheets["Problemas"].columns) == ["fila", "ncf", "severidad", "regla", "mensaje"]
    assert set(sheets["Facturas"]["estado"]) == {"OK", "ADVERTENCIA", "ERROR"}
    # the example triggers every rule
    assert set(sheets["Problemas"]["regla"]) == set(ci.RULES)
    out = capsys.readouterr().out
    assert "Facturas revisadas" in out and "Reporte:" in out


def test_excel_input_with_numeric_ids_and_dates(tmp_path):
    src = tmp_path / "facturas.xlsx"
    pd.DataFrame([{
        "Fecha": dt.datetime(2026, 9, 15), "RNC Emisor": 101850043, "RNC Comprador": 113918205,
        "NCF": "B0100000001", "Cantidad": 1, "Precio Unitario": 1000, "Tasa ITBIS": 18,
        "ITBIS": 180, "Total": 1180,
    }]).to_excel(src, index=False)
    before = sha(src)
    assert ci.main([str(src), "--today", "2026-10-07"]) == 0
    assert sha(src) == before
    problems = pd.read_excel(tmp_path / "facturas_reporte.xlsx", sheet_name="Problemas")
    assert list(problems["regla"]) == ["R1-CEDULA-CEROS"]


def test_unknown_columns_exit_code_2(tmp_path, capsys):
    src = tmp_path / "raro.csv"
    src.write_text("a;b;c\n1;2;3\n", encoding="utf-8")
    assert ci.main([str(src)]) == 2
    assert "--map" in capsys.readouterr().err
    assert not (tmp_path / "raro_reporte.xlsx").exists()


def test_missing_file_exit_code_1(tmp_path):
    assert ci.main([str(tmp_path / "no_existe.csv")]) == 1


def test_semicolon_cp1252_csv(tmp_path):
    src = tmp_path / "excel_es.csv"
    text = ("Fecha;RNC Emisor;RNC/Cédula Comprador;NCF;Descripción;Cantidad;Precio Unitario;"
            "Tasa ITBIS;ITBIS;Total\n15/09/2026;101850043;130123454;B0100000001;Café;1;1.000,00;"
            "18;180,00;1.180,00\n")
    src.write_bytes(text.encode("cp1252"))
    assert ci.main([str(src), "--today", "2026-10-07"]) == 0
    problems = pd.read_excel(tmp_path / "excel_es_reporte.xlsx", sheet_name="Problemas")
    assert problems.empty


def test_ley_30_26_note_always_shown_in_spanish():
    summary = ci.format_summary(run(line()), Path("x.csv"), None, RULES, TODAY)
    assert "NOTA: La Ley 30-26 (junio 2026) agregó nuevas exenciones" in summary
    assert "no se ha confirmado contra el texto de la ley" in summary
    assert "confírmalo con tu contador" in summary
