[Español](README.md) · **English**

![DR Facturación: check your invoices before you file](docs/banner.en.png)

<p align="center"><strong>Check the RNC, NCF/e-NCF and ITBIS on your invoices in seconds, and get an Excel report with every issue explained in plain Spanish.</strong></p>

## What is this?

You give it your invoices in an Excel file. It tells you which ones have errors and how to fix them.
It never changes your file: it only reads it and creates a separate new report.

## Who is it for?

Business owners, accountants and freelancers in the Dominican Republic who want to catch
errors in their invoices before filing or sending them to their accountant.

If you are a small or micro taxpayer, you have until **November 15, 2026** to start issuing
electronic invoices (e-CF).

> [!IMPORTANT]
> **It doesn't change your file.** It only reads it and creates a new report.
> **Your data never leaves your computer.** It doesn't connect to DGII or anywhere else.
> **It isn't tax advice.** Review the results with your accountant before filing.

## How to use it (the easy way)

1. **Put your invoices in Excel.**
   [Download the template](https://github.com/alejandrocimentada/dr-facturacion/raw/main/templates/plantilla_facturas.xlsx)
   and fill it in, with one row for each product or service on each invoice.
   You can also use your own Excel file if it has similar columns (date, RNC, NCF, quantity,
   price, ITBIS, total).

2. **Ask your AI assistant to check it.** If you use Claude Code or another AI agent
   (an assistant that can work with files on your computer), copy and paste this message:

   ```text
   Instala el skill dr-facturacion desde https://github.com/alejandrocimentada/dr-facturacion
   y luego revisa mi archivo facturas.xlsx
   ```

   (In English: "Install the dr-facturacion skill from https://github.com/alejandrocimentada/dr-facturacion
   and then check my file facturas.xlsx".) The agent does the whole installation for you.
   Replace `facturas.xlsx` with the name of your file.

3. **Open the report and fix.** The agent creates a new file called `facturas_reporte.xlsx`
   next to your file. Open it and fix the rows marked in red (errors). Rows in orange are
   warnings: check those too.

The report has three tabs:

**Resumen** (Summary): how many invoices were checked, how many issues there are, and whether the ITBIS adds up.

![Resumen tab](docs/reporte-resumen.png)

**Problemas** (Issues): one issue per row, with what's wrong and what to do. The most serious come first.

![Problemas tab](docs/reporte-problemas.png)

**Facturas** (Invoices): your data exactly as it was, with each row colored by status (red, orange or green).

![Facturas tab](docs/reporte-facturas.png)

## What does each error mean?

| Error | What it means | What to do |
|---|---|---|
| The issuer's RNC or cédula isn't valid | There's probably a typo, or it's missing. | Copy the exact RNC from an invoice or official document. |
| The buyer's RNC or cédula isn't valid | There's probably a typo. | Ask the customer for the correct RNC or cédula. |
| Cédula missing its leading zeros | Excel deleted the zeros at the start. The tool filled them back in for you. | Nothing urgent. To avoid it, format that column as text. |
| The NCF has the wrong format | The receipt number has too many or too few characters. | Copy it again from the original invoice. |
| Unknown receipt type | The letters and numbers at the start of the NCF (like B01) don't exist. | Check that the NCF is typed correctly. |
| Crédito fiscal without the buyer's RNC | A B01/E31 invoice must always include the buyer's RNC. | Add the customer's RNC, or use a consumer invoice (B02). |
| Consumer invoice to a company | You gave a B02/E32 invoice to someone who has an RNC. | Check with the customer: they may have needed crédito fiscal (B01). |
| ITBIS on an invoice that shouldn't have it | Special-regime or export invoices almost never carry ITBIS. | Ask your accountant whether that ITBIS belongs there. |
| Credit or debit note without the original invoice | The note doesn't say which invoice it corrects. | Enter the original invoice's NCF in the `ncf_modificado` column. |
| Repeated NCF | The same receipt number appears on two different invoices. | Find out which one is right. Each invoice needs its own NCF. |
| Repeated line | The same line appears twice on one invoice. | Delete it if it was copied by mistake. |
| Gap in the numbering | Some numbers are missing between one NCF and the next. | They may be voided invoices. If so, they belong in DGII's voided-receipts report (formato 608). |
| Expired sequence | The invoice was issued after those NCFs expired. | Request a new sequence from DGII and talk to your accountant. |
| Invalid ITBIS rate | The rate isn't 18%, 16%, 0% or "exento" (exempt). | Fix the rate. The usual one is 18%. |
| Empty or unreadable number | An amount is missing or has letters where numbers go. | Fill in the quantity, price, ITBIS and total. |
| The ITBIS doesn't add up | The ITBIS isn't quantity × price × rate. | Recalculate the ITBIS for that line. |
| The total doesn't add up | The invoice total isn't the amount before ITBIS + the ITBIS. | Recalculate the total. |
| Line at the 16% rate | Only some products carry 16%. This is a notice, not an error. | Confirm that the product really carries 16%. |
| Empty or invalid date | The date is missing, or doesn't exist (for example, February 31). | Enter the correct date (DD/MM/YYYY). |
| Future date | The date is after today. | Fix the year, month or day. |

## Glossary

- **RNC**: Registro Nacional de Contribuyentes. The 9-digit number DGII uses to identify a company.
- **Cédula**: a person's 11-digit national ID number. It works as the RNC for individuals.
- **NCF**: Número de Comprobante Fiscal. The official number of each paper invoice, like `B0100000001`.
- **e-NCF / e-CF**: the electronic version. The e-CF is the electronic invoice; the e-NCF is its number, like `E310000000001`.
- **ITBIS**: the tax charged on sales of goods and services (similar to VAT). Usually 18%.
- **Crédito fiscal**: the ITBIS a company paid on its purchases, which it can subtract from the ITBIS it owes.
- **B01 vs B02**: B01 is a crédito fiscal invoice (for companies, includes the buyer's RNC). B02 is a consumer invoice (for individuals, no RNC).

## FAQ

**Is my data sent anywhere?**
No. Everything is checked on your computer. It doesn't connect to DGII or the internet.

**Does it replace my accountant?**
No. It helps you find errors before you send your invoices. Your accountant has the final word.

**Can it read PDF invoices?**
Not yet. For now it only reads Excel (.xlsx) and CSV (a table format Excel can save).

**Does it check whether an RNC is active?**
No. It only checks that the number is well-formed. To see if it's active, look it up on DGII's website.

**Do I need to know how to code?**
No, if you use an AI agent like Claude Code. The agent installs and runs it for you.

---

## For technical users

**dr-facturacion** is a skill for AI agents (Claude Code, Codex and other agents that
support `SKILL.md`) that also works as a standalone command. The November 15, 2026 e-CF
deadline comes from DGII Aviso 06-26
([summary in Spanish](https://siemprealdia.co/republica-dominicana/impuestos/dgii-prorrogo-el-e-cf-para-pequenos-micros-y-no-clasificados/)).

### Manual installation

Requirements: Python 3.10 or newer and `pip`.

**Claude Code (macOS / Linux)**

```bash
git clone https://github.com/alejandrocimentada/dr-facturacion.git
mkdir -p ~/.claude/skills/dr-facturacion
cp -R dr-facturacion/. ~/.claude/skills/dr-facturacion/
pip install -r ~/.claude/skills/dr-facturacion/requirements.txt
```

**Claude Code (Windows PowerShell)**

```powershell
git clone https://github.com/alejandrocimentada/dr-facturacion.git
New-Item -ItemType Directory -Force "$HOME\.claude\skills\dr-facturacion" | Out-Null
Copy-Item -Recurse -Force "dr-facturacion\*" "$HOME\.claude\skills\dr-facturacion\"
pip install -r "$HOME\.claude\skills\dr-facturacion\requirements.txt"
```

**Codex**

```bash
git clone https://github.com/alejandrocimentada/dr-facturacion.git
mkdir -p ~/.codex/skills/dr-facturacion
cp -R dr-facturacion/. ~/.codex/skills/dr-facturacion/
pip install -r ~/.codex/skills/dr-facturacion/requirements.txt
```

Once installed, ask the agent in a conversation:

> Use dr-facturacion to check my September invoices: facturas_septiembre.xlsx

The agent runs the check, explains the main issues grouped by severity with row numbers,
and tells you where the report is.

### Command-line use (no agent)

```bash
git clone https://github.com/alejandrocimentada/dr-facturacion.git
cd dr-facturacion
pip install -r requirements.txt
python scripts/check_invoices.py facturas.xlsx
```

The report is saved next to the file as `facturas_reporte.xlsx`. Other options:
`--sheet "Septiembre"` (Excel sheet), `--output revision.xlsx` (report path) and
`--map ncf="No. Comprobante"` (map a column that isn't recognized).

### Columns and mapping (`--map`)

Use [`templates/plantilla_facturas.xlsx`](templates/plantilla_facturas.xlsx): one row per
invoice line. Headers may be in Spanish or English (case, accents and punctuation are
ignored), and CSV or Excel both work. If a column has a different name, map it with
`--map field="Your column name"`.

| Column | Required | Description |
|---|---|---|
| `fecha` | yes | issue date (DD/MM/YYYY) |
| `rnc_emisor` | yes | issuer RNC or cédula |
| `rnc_comprador` | yes (may be empty for consumers) | buyer RNC or cédula |
| `ncf` | yes | NCF or e-NCF |
| `descripcion` | no | product or service |
| `cantidad` | yes | quantity |
| `precio_unitario` | yes | unit price **without** ITBIS |
| `tasa_itbis` | yes | `18`, `16`, `0` or `exento` |
| `itbis` | yes | ITBIS charged on the line |
| `total` | yes | line total (base + ITBIS) |
| `vencimiento_secuencia` | no | NCF sequence expiry date |
| `ncf_modificado` | no | original NCF, for debit/credit notes |

The report screenshots above were generated from
[`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv).

### What does it check?

| Check | What it catches | Severity |
|---|---|---|
| RNC / cédula | issuer or buyer RNC (9 digits) or cédula (11 digits) missing, malformed or with an invalid check digit; cédulas that lost their leading zeros in Excel | ERROR · INFO |
| NCF / e-NCF structure | NCF not following `B` + type + 8 digits, e-NCF not following `E` + type + 10 digits, receipt types that don't exist, pre-2018 format | ERROR |
| NCF type vs buyer | crédito fiscal (B01/E31) without a buyer RNC; consumo (B02/E32) issued to a company with an RNC; special regime or export (B14/E44, B16/E46) with ITBIS; debit/credit notes without the modified NCF | ERROR · ADVERTENCIA |
| Sequences | duplicate NCFs, repeated lines, expired sequences, numbering gaps (possible voided receipts for formato 608) | ERROR · ADVERTENCIA · INFO |
| ITBIS math | rates other than 18%, 16%, 0% or exempt; line ITBIS different from quantity × price × rate; unreadable amounts; 16% lines to confirm | ERROR · INFO |
| Totals | invoice total different from base + ITBIS | ERROR |
| Dates | missing, impossible (e.g. 31/02) or future dates | ERROR |

Each of the 20 rules is defined in [`scripts/check_invoices.py`](scripts/check_invoices.py) (the `RULES` dictionary).

### How well has it been verified?

- **Rules checked against DGII sources on 2026-10-07** (the `last_verified` field in
  [`rules/itbis_rates.yaml`](rules/itbis_rates.yaml) and [`rules/ncf_types.yaml`](rules/ncf_types.yaml)):
  - [DGII: ITBIS](https://dgii.gov.do/cicloContribuyente/obligacionesTributarias/principalesImpuestos/Paginas/Itbis.aspx)
  - [DGII: Guía #7 ITBIS (PDF)](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/itbis/Documents/1-Guia%207%20-%20%28ITBIS%29.pdf)
  - [DGII: Types of fiscal receipts](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/tiposComprobantes.aspx)
  - [DGII: Guía Informativa sobre Comprobantes Fiscales (PDF)](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/2-Guia-Informativa-NCF.pdf)
  - [DGII: Guía de Comprobantes Fiscales Especiales (PDF)](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/3-Guia-Comprobantes-Fiscales-Especiales-NG-05-19.pdf)
  - [DGII: Electronic fiscal receipts (e-CF)](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/comprobantesFiscalesElectronicos.aspx)
  - [DGII: Aviso 24-19, e-NCF structure (PDF)](https://dgii.gov.do/publicacionesOficiales/avisosInformativos/Documents/2019/24-19.pdf)
  - Check digits: [python-stdnum `stdnum.do.rnc`](https://arthurdejong.org/python-stdnum/doc/2.2/stdnum.do.rnc.html) and [`stdnum.do.cedula`](https://arthurdejong.org/python-stdnum/doc/2.2/stdnum.do.cedula.html)
- **140 automated tests**, with valid and invalid cases for every rule.
- **Tested only with the sample file** ([`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv)),
  not yet with real company invoices.
- The **sequence-gap and repeated-line** rules and the rounding **tolerances**
  (RD$0.01 per line, RD$0.05 per invoice) are this tool's own heuristics, not DGII rules.
- **Ley 30-26:** Ley 30-26 (June 2026) added new ITBIS exemptions for selected goods, effective immediately. Published summaries of the reform don't list changes to the 18% / 16% rates, but this hasn't been confirmed against the law's text. This tool doesn't check product lists: if a product may now be exempt, or if you're unsure about a rate, confirm with your accountant.
  ([Source: RSM, Law 30-26](https://www.rsm.global/dominicanrepublic/en/news/dominican-republic-tax-reform-law-30-26))

### Why can you trust it?

1. **It never touches the input file.** It reads it into memory and writes a separate
   report; the tests verify with a hash that the file stays identical (CSV and Excel).
2. **Check digits are validated with [python-stdnum](https://arthurdejong.org/python-stdnum/)**,
   an open, maintained library, not home-made formulas.
3. **Rules live in editable YAML files** ([`rules/`](rules/)), with their verification date
   and sources. If the rates haven't been verified in more than 90 days, the tool warns you.
4. **Every issue explains what's wrong and what to do**, in plain Spanish, with the row number and NCF.

### Customizing the rules

The rules live in two YAML files you can edit without touching the code:

- [`rules/itbis_rates.yaml`](rules/itbis_rates.yaml): accepted ITBIS rates (`general`,
  `reducida`, `exportacion`, `exento`), the Ley 30-26 note (`note` / `nota_es`) and their sources.
- [`rules/ncf_types.yaml`](rules/ncf_types.yaml): series (`B`, `E`) with their length, and the valid
  receipt types with their flags (`requires_buyer_id`, `consumer`, `usually_exempt`, `modifies`).

When you change something, **update `last_verified`** to today's date and **add the source**
under `sources`, so users know where each rule comes from.

### Development and validation

```bash
pip install -r requirements.txt pytest
python -m pytest                                          # 140 tests
python scripts/check_invoices.py examples/facturas_ejemplo.csv
python scripts/make_template.py                           # rebuild the Excel template
```

The sample file triggers all 20 rules at least once. The banner is generated from
[`docs/banner.html`](docs/banner.html).

### Known limitations

- **No online DGII lookup**: it can't confirm that an RNC is active or that an NCF was
  authorized, only that they are well-formed.
- **No product lists**: it doesn't know which products qualify for the 16% rate or are exempt.
- **No PDF support yet**: CSV and Excel (.xlsx) only.
- **No withholdings**: ITBIS and ISR withholdings are not covered (nor ISC or the legal tip).
- **Rules can change**: check the `last_verified` date in [`rules/`](rules/) before relying on the results.

## License

[MIT](LICENSE) © 2026 Alejandro Cimentada.
