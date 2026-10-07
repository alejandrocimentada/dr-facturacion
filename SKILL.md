---
name: dr-facturacion
description: Check, validate or audit Dominican Republic invoices (facturas, comprobantes fiscales) from a CSV or Excel file - RNC/cédula check digits, NCF/e-NCF structure and type, NCF type vs buyer, duplicate/expired sequences and ITBIS math - and produce an Excel error report. Use when the user wants to review facturas, NCF, e-CF, ITBIS or comprobantes before filing (606/607) or sending them to their accountant.
---

# dr-facturacion

Checks Dominican invoice lines and writes `<input-name>_reporte.xlsx` next to the input
(sheets: Resumen, Problemas, Facturas) plus a short Spanish summary on stdout.
All paths below are relative to this skill's directory.

## Steps

1. **Find the file.** Use the path the user gave; otherwise look for `.csv`/`.xlsx` files
   in the working directory and ask which one if there is more than one. If the user has
   no file yet, point them to `templates/plantilla_facturas.xlsx`.
2. **Check requirements** with `python -c "import pandas, openpyxl, stdnum, yaml"`.
   If it fails, run `python -m pip install -r requirements.txt` (Python 3.10+).
3. **Run** `python scripts/check_invoices.py "<file>"` (add `--sheet "<name>"` for a
   specific Excel sheet).
   - Exit code 2 = columns not recognized. The error lists the missing fields and the
     headers found. Show both to the user, ask which header matches each missing field,
     then re-run with `--map campo="Encabezado"` (repeatable). Do not guess silently.
   - Exit code 1 = file missing/unreadable: tell the user.
4. **Read the summary** printed by the script.
5. **Explain the results** in plain Spanish (or the user's language): group by severity
   (ERROR first, then ADVERTENCIA, then INFO), cite row numbers (`fila`) and NCF, say what
   is wrong and what to do. Summarize repeated issues instead of listing every row.
6. **Point to the report**: give the path of the `_reporte.xlsx` file and mention the
   Problemas sheet (all issues) and Facturas sheet (each line colored OK/ADVERTENCIA/ERROR).

## Rules

- Never edit, rename, move or "fix" the invoice file. The script only reads it; if the
  user wants corrections, tell them what to change and let them do it.
- Never present results as tax advice. They are automated format and arithmetic checks;
  the script cannot confirm with DGII that an RNC is active or an NCF was authorized.
- If the summary contains an `AVISO` about ITBIS rates (last verified more than 90 days
  ago), warn the user that the rates may be outdated and must be confirmed at dgii.gov.do.
- Pass on the `NOTA` about Ley 30-26 printed in the summary: new ITBIS exemptions; no
  rate changes in published summaries, but not confirmed against the law's text.
- Always end your answer with exactly:
  "Revisa estos resultados con tu contador antes de presentar cualquier declaración."

## Severity meaning (for explanations)

- **ERROR**: the comprobante or amount is wrong as recorded and should be corrected.
- **ADVERTENCIA**: probably a problem; needs a human decision.
- **INFO**: worth confirming (sequence gaps, 16% rate, cédula missing leading zeros).
