# dr-facturacion

**[English](#english) · [Español](#español)**

---

## English

An agent skill (Claude Code and other agents that support `SKILL.md`) and a standalone
CLI that checks Dominican Republic invoices — RNC/cédula, NCF/e-NCF and ITBIS — and
produces an Excel error report.

> **Not tax advice.** This tool runs automated format and arithmetic checks. Review the
> results with your accountant before filing any return.

### What it does

Given a CSV or Excel file with one row per invoice line, it:

1. validates the issuer and buyer **RNC (9 digits) / cédula (11 digits)** check digits;
2. validates the **NCF** (`B` + type + 8 digits) / **e-NCF** (`E` + type + 10 digits) structure and type;
3. checks the **NCF type against the buyer** (crédito fiscal without buyer ID, consumo issued to a company, etc.);
4. finds **duplicate NCFs, sequence gaps and expired sequences**;
5. recalculates **ITBIS** per line and the **invoice total**;
6. rejects **unreadable or future dates**.

It writes `<input-name>_reporte.xlsx` next to the input with three sheets:

| Sheet | Content |
|---|---|
| **Resumen** | invoices and lines checked, ERROR / ADVERTENCIA / INFO counts, ITBIS charged vs recalculated and the difference, when the rates were last verified |
| **Problemas** | one row per issue: `fila`, `ncf`, `severidad`, `regla`, `mensaje` (plain Spanish: what's wrong and what to do) |
| **Facturas** | the original data plus an `estado` column (OK / ADVERTENCIA / ERROR), color-coded |

and prints a short Spanish summary to the terminal. **The input file is never modified**:
it is read into memory and the report is a new file (if a report with that name is open in
Excel, a numbered copy is written instead).

### What the report looks like

Report generated from [`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv):

**Resumen**: totals, severity counts, ITBIS charged vs recalculated

![Resumen sheet](docs/reporte-resumen.png)

**Problemas**: one row per issue, sorted by severity

![Problemas sheet](docs/reporte-problemas.png)

**Facturas**: original data with a color-coded `estado` column

![Facturas sheet](docs/reporte-facturas.png)

### Install

Requires Python 3.10+.

**As a Claude Code skill**

```bash
# personal skill (all projects)
git clone https://github.com/<you>/dr-facturacion ~/.claude/skills/dr-facturacion
# or project skill
git clone https://github.com/<you>/dr-facturacion .claude/skills/dr-facturacion

python -m pip install -r ~/.claude/skills/dr-facturacion/requirements.txt
```

Then just ask: *"revisa las facturas en ventas_septiembre.xlsx"*. The agent follows
[`SKILL.md`](SKILL.md): it runs the checker, explains the top issues by severity with row
numbers, and points you to the report. (If you skip the `pip install`, the agent installs
the requirements the first time.)

**As a standalone CLI**

```bash
git clone https://github.com/<you>/dr-facturacion
cd dr-facturacion
python -m pip install -r requirements.txt
python scripts/check_invoices.py examples/facturas_ejemplo.csv
```

### Usage

```bash
python scripts/check_invoices.py facturas.xlsx
python scripts/check_invoices.py facturas.xlsx --sheet "Septiembre"
python scripts/check_invoices.py ventas.csv --output revision.xlsx
# headers the checker does not recognize: map them by hand (repeatable)
python scripts/check_invoices.py ventas.csv --map fecha="Fecha Doc" --map ncf="No. Comprobante"
```

Exit codes: `0` checked (even if problems were found) · `1` file missing/unreadable ·
`2` columns not recognized (use `--map`).

CSV files may use `,` `;` tab or `|` as separator and be UTF-8 or Windows-1252 encoded.
Amounts like `1,180.00`, `1.180,00` and `RD$ 1,180.00` are understood. Dates are
day-first (`DD/MM/AAAA`), ISO (`AAAA-MM-DD`) or real Excel dates.

### Column template

Use [`templates/plantilla_facturas.xlsx`](templates/plantilla_facturas.xlsx). Headers may be in
Spanish or English and are matched ignoring case, accents and punctuation (e.g.
`RNC/Cédula Comprador`, `Buyer RNC`, `ITBIS %`).

| Column | Required | Description |
|---|---|---|
| `fecha` | yes | issue date |
| `rnc_emisor` | yes | issuer RNC or cédula |
| `rnc_comprador` | yes (may be empty for consumers) | buyer RNC or cédula |
| `ncf` | yes | NCF or e-NCF |
| `descripcion` | no | item description |
| `cantidad` | yes | quantity |
| `precio_unitario` | yes | unit price **without** ITBIS |
| `tasa_itbis` | yes | `18`, `16`, `0` or `exento` (also `18%`, `0.18`) |
| `itbis` | yes | ITBIS charged on the line (blank = 0) |
| `total` | yes | line total (base + ITBIS). A per-invoice total repeated on every line is also accepted |
| `vencimiento_secuencia` | no | NCF sequence expiry date |
| `ncf_modificado` | no | NCF modified by a credit/debit note |

Lines with the same issuer + NCF + date + buyer are treated as one invoice.

### How has it been verified?

Rules and rates were last verified on **2026-10-07** (`last_verified` in
[`rules/itbis_rates.yaml`](rules/itbis_rates.yaml) and [`rules/ncf_types.yaml`](rules/ncf_types.yaml)).
The checker prints a warning when the rates are more than 90 days old. Every rule has
valid and invalid test cases in [`tests/`](tests/) (`pytest`), and
[`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv) triggers every rule at least once.

| Rule | Severity | Check | Source |
|---|---|---|---|
| R1-RNC-EMISOR | ERROR | issuer RNC/cédula present, 9 or 11 digits, valid check digit | python-stdnum [`stdnum.do.rnc`](https://arthurdejong.org/python-stdnum/doc/latest/stdnum.do.rnc.html), [`stdnum.do.cedula`](https://arthurdejong.org/python-stdnum/doc/latest/stdnum.do.cedula.html) |
| R1-RNC-COMPRADOR | ERROR | same for the buyer (when present) | same |
| R1-CEDULA-CEROS | INFO | 9/10-digit number that is only valid as a cédula once the leading zeros are restored (Excel number formatting) | same |
| R2-NCF-ESTRUCTURA | ERROR | `B`+2+8 = 11 chars or `E`+2+10 = 13 chars; pre-2018 `A…`/`P…` format rejected | DGII [Tipos de comprobantes](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/tiposComprobantes.aspx), [Aviso 24-19 (e-NCF structure)](https://dgii.gov.do/publicacionesOficiales/avisosInformativos/Documents/2019/24-19.pdf) |
| R2-NCF-TIPO | ERROR | type listed in `rules/ncf_types.yaml` (B01–04, 11–17; E31–34, 41, 43–47) | DGII [Guía Informativa NCF](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/2-Guia-Informativa-NCF.pdf), [e-CF](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/comprobantesFiscalesElectronicos.aspx) |
| R3-CREDITO-SIN-COMPRADOR | ERROR | B01/E31 needs a valid buyer RNC/cédula | DGII Guía Informativa NCF |
| R3-CONSUMO-CON-RNC | ADVERTENCIA | B02/E32 issued to a buyer with an RNC (they can't use it for ITBIS credit or ISR expense) | DGII Guía Informativa NCF |
| R3-EXENTO-CON-ITBIS | ADVERTENCIA | B14/E44 or B16/E46 with ITBIS charged (normally exempt / 0%) | DGII [Guía Comprobantes Especiales](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/3-Guia-Comprobantes-Fiscales-Especiales-NG-05-19.pdf) |
| R3-NOTA-SIN-REFERENCIA | ADVERTENCIA | B03/B04/E33/E34 without `ncf_modificado` | DGII Guía Informativa NCF |
| R4-NCF-DUPLICADO | ERROR | same issuer + NCF used for invoices with different date or buyer | DGII Guía Informativa NCF (each NCF is unique) |
| R4-LINEA-REPETIDA | ADVERTENCIA | identical line repeated inside one invoice (possible double entry) | heuristic of this tool |
| R4-SECUENCIA-SALTO | INFO | gaps in the sequence per issuer + type (may be voided; report in formato 608) | heuristic of this tool |
| R4-SECUENCIA-VENCIDA | ERROR | `fecha` after `vencimiento_secuencia` | DGII Tipos de comprobantes |
| R5-TASA-INVALIDA | ERROR | rate is one of `rules/itbis_rates.yaml`: 18%, 16%, 0%, exento | DGII [ITBIS](https://dgii.gov.do/cicloContribuyente/obligacionesTributarias/principalesImpuestos/Paginas/Itbis.aspx), [Guía #7 ITBIS](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/itbis/Documents/1-Guia%207%20-%20(ITBIS).pdf) |
| R5-NUMERO-INVALIDO | ERROR | quantity, price, ITBIS or total missing or unreadable | input sanity check |
| R5-ITBIS-LINEA | ERROR | `cantidad × precio_unitario × tasa` vs ITBIS charged, tolerance RD$0.01 per line | DGII ITBIS, Guía #7 (tolerance is a choice of this tool) |
| R5-TOTAL-FACTURA | ERROR | Σ total vs Σ base + Σ ITBIS, tolerance RD$0.05 per invoice | arithmetic (tolerance is a choice of this tool) |
| R5-TASA-REDUCIDA | INFO | any 16% line: confirm the product is on the reduced-rate list | DGII Guía #7; see the Ley 30-26 note in `rules/itbis_rates.yaml` |
| R6-FECHA-INVALIDA | ERROR | date missing or not a real date | input sanity check |
| R6-FECHA-FUTURA | ERROR | date after today | input sanity check |

**About python-stdnum and NCF types.** The checker also calls `stdnum.do.ncf`. If stdnum
rejects a type that is listed in `rules/ncf_types.yaml` (stdnum can lag behind DGII changes),
the YAML wins and the terminal summary notes it. As of python-stdnum 2.2, both lists match.

**Ley 30-26.** Ley 30-26 (June 2026) added new ITBIS exemptions for selected goods, effective immediately. Published summaries of the reform don't list changes to the 18% / 16% rates, but this hasn't been confirmed against the law's text. This tool doesn't check product lists: if a product may now be exempt, or if you're unsure about a rate, confirm with your accountant. Source: [RSM, Law 30-26](https://www.rsm.global/dominicanrepublic/en/news/dominican-republic-tax-reform-law-30-26).

### Limitations

- **No online DGII lookup.** It can't tell whether an RNC is active, whether an NCF was
  authorized for that issuer, or whether an e-CF was accepted — only whether they are well-formed.
- It doesn't know which products qualify for 16% or exemption; it flags 16% lines for review.
- B02/E32 warnings only fire for 9-digit RNCs, not for cédulas.
- Only `B` and `E` series are accepted (DGII allows other e-CF series letters in some cases).
- Prices are assumed to exclude ITBIS. ISC, tips (propina legal) and other taxes are not checked.
- Duplicate detection relies on date/buyer differences; an invoice pasted twice with identical
  lines is reported as repeated lines (ADVERTENCIA), not as a duplicate NCF.

### Development

```bash
python -m pip install -r requirements.txt pytest
python -m pytest
python scripts/make_template.py   # regenerate the Excel template
```

### Disclaimer

This project is not tax advice and is not affiliated with the DGII. Rates and rules change;
always confirm with the DGII and review the results with your accountant before filing any return.

License: [MIT](LICENSE).

---

## Español

Una skill para agentes (Claude Code y otros agentes compatibles con `SKILL.md`) y un
programa de línea de comandos que revisa facturas dominicanas — RNC/cédula, NCF/e-NCF e
ITBIS — y genera un reporte de errores en Excel.

> **No es asesoría fiscal.** La herramienta hace verificaciones automáticas de formato y
> aritmética. Revisa los resultados con tu contador antes de presentar cualquier declaración.

### Qué hace

A partir de un CSV o Excel con una fila por línea de factura:

1. valida el dígito verificador del **RNC (9 dígitos) / cédula (11 dígitos)** del emisor y del comprador;
2. valida la estructura y el tipo del **NCF** (`B` + tipo + 8 dígitos) / **e-NCF** (`E` + tipo + 10 dígitos);
3. revisa el **tipo de NCF contra el comprador** (crédito fiscal sin RNC del comprador, consumo emitido a una empresa, etc.);
4. detecta **NCF duplicados, saltos de secuencia y secuencias vencidas**;
5. recalcula el **ITBIS** de cada línea y el **total de la factura**;
6. rechaza **fechas ilegibles o futuras**.

Genera `<nombre-del-archivo>_reporte.xlsx` junto al archivo de entrada, con tres hojas:

| Hoja | Contenido |
|---|---|
| **Resumen** | facturas y líneas revisadas, cantidad de ERROR / ADVERTENCIA / INFO, ITBIS cobrado vs recalculado y la diferencia, fecha de verificación de las tasas |
| **Problemas** | una fila por problema: `fila`, `ncf`, `severidad`, `regla`, `mensaje` (en español sencillo: qué está mal y qué hacer) |
| **Facturas** | los datos originales más una columna `estado` (OK / ADVERTENCIA / ERROR) con colores |

y muestra un resumen corto en español en la terminal. **El archivo de entrada nunca se
modifica**: se lee en memoria y el reporte es un archivo nuevo (si un reporte con ese nombre
está abierto en Excel, se guarda una copia numerada).

### Cómo se ve el reporte

Reporte generado a partir de [`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv):

**Resumen**: totales, cantidad por severidad, ITBIS cobrado vs recalculado

![Hoja Resumen](docs/reporte-resumen.png)

**Problemas**: una fila por problema, ordenadas por severidad

![Hoja Problemas](docs/reporte-problemas.png)

**Facturas**: datos originales con la columna `estado` en colores

![Hoja Facturas](docs/reporte-facturas.png)

### Instalación

Requiere Python 3.10+.

**Como skill de Claude Code**

```bash
# skill personal (todos los proyectos)
git clone https://github.com/<tu-usuario>/dr-facturacion ~/.claude/skills/dr-facturacion
# o skill del proyecto
git clone https://github.com/<tu-usuario>/dr-facturacion .claude/skills/dr-facturacion

python -m pip install -r ~/.claude/skills/dr-facturacion/requirements.txt
```

Luego pide: *"revisa las facturas en ventas_septiembre.xlsx"*. El agente sigue
[`SKILL.md`](SKILL.md): ejecuta el verificador, explica los problemas principales por
severidad con números de fila y te indica dónde está el reporte. (Si no instalas los
requisitos, el agente los instala la primera vez.)

**Como programa independiente**

```bash
git clone https://github.com/<tu-usuario>/dr-facturacion
cd dr-facturacion
python -m pip install -r requirements.txt
python scripts/check_invoices.py examples/facturas_ejemplo.csv
```

### Uso

```bash
python scripts/check_invoices.py facturas.xlsx
python scripts/check_invoices.py facturas.xlsx --sheet "Septiembre"
python scripts/check_invoices.py ventas.csv --output revision.xlsx
# encabezados que no se reconocen: asígnalos a mano (repetible)
python scripts/check_invoices.py ventas.csv --map fecha="Fecha Doc" --map ncf="No. Comprobante"
```

Códigos de salida: `0` revisado (aunque haya problemas) · `1` archivo inexistente o
ilegible · `2` columnas no reconocidas (usa `--map`).

Los CSV pueden usar `,` `;` tabulador o `|` como separador y estar en UTF-8 o Windows-1252.
Se entienden montos como `1,180.00`, `1.180,00` y `RD$ 1,180.00`. Las fechas pueden ser
día primero (`DD/MM/AAAA`), ISO (`AAAA-MM-DD`) o fechas reales de Excel.

### Plantilla de columnas

Usa [`templates/plantilla_facturas.xlsx`](templates/plantilla_facturas.xlsx). Los encabezados
pueden estar en español o inglés y se comparan sin importar mayúsculas, acentos ni signos
(p. ej. `RNC/Cédula Comprador`, `Buyer RNC`, `ITBIS %`).

| Columna | Obligatoria | Descripción |
|---|---|---|
| `fecha` | sí | fecha de emisión |
| `rnc_emisor` | sí | RNC o cédula del emisor |
| `rnc_comprador` | sí (puede ir vacía en consumo) | RNC o cédula del comprador |
| `ncf` | sí | NCF o e-NCF |
| `descripcion` | no | descripción del artículo |
| `cantidad` | sí | cantidad |
| `precio_unitario` | sí | precio unitario **sin** ITBIS |
| `tasa_itbis` | sí | `18`, `16`, `0` o `exento` (también `18%`, `0.18`) |
| `itbis` | sí | ITBIS cobrado en la línea (vacío = 0) |
| `total` | sí | total de la línea (base + ITBIS). También se acepta el total de la factura repetido en cada línea |
| `vencimiento_secuencia` | no | fecha de vencimiento de la secuencia de NCF |
| `ncf_modificado` | no | NCF que modifica una nota de crédito/débito |

Las líneas con el mismo emisor + NCF + fecha + comprador se consideran una sola factura.

### ¿Cómo se ha verificado?

Las reglas y tasas se verificaron por última vez el **2026-10-07** (`last_verified` en
[`rules/itbis_rates.yaml`](rules/itbis_rates.yaml) y [`rules/ncf_types.yaml`](rules/ncf_types.yaml)).
El verificador avisa cuando las tasas tienen más de 90 días sin verificarse. Cada regla tiene
casos válidos e inválidos en [`tests/`](tests/) (`pytest`), y
[`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv) activa todas las reglas al menos una vez.

| Regla | Severidad | Verificación | Fuente |
|---|---|---|---|
| R1-RNC-EMISOR | ERROR | RNC/cédula del emisor presente, 9 u 11 dígitos, dígito verificador válido | python-stdnum [`stdnum.do.rnc`](https://arthurdejong.org/python-stdnum/doc/latest/stdnum.do.rnc.html), [`stdnum.do.cedula`](https://arthurdejong.org/python-stdnum/doc/latest/stdnum.do.cedula.html) |
| R1-RNC-COMPRADOR | ERROR | lo mismo para el comprador (si existe) | igual |
| R1-CEDULA-CEROS | INFO | número de 9/10 dígitos que solo es válido como cédula al restaurar los ceros iniciales (formato numérico de Excel) | igual |
| R2-NCF-ESTRUCTURA | ERROR | `B`+2+8 = 11 caracteres o `E`+2+10 = 13; se rechaza el formato anterior a 2018 `A…`/`P…` | DGII [Tipos de comprobantes](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/tiposComprobantes.aspx), [Aviso 24-19 (estructura e-NCF)](https://dgii.gov.do/publicacionesOficiales/avisosInformativos/Documents/2019/24-19.pdf) |
| R2-NCF-TIPO | ERROR | tipo listado en `rules/ncf_types.yaml` (B01–04, 11–17; E31–34, 41, 43–47) | DGII [Guía Informativa NCF](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/2-Guia-Informativa-NCF.pdf), [e-CF](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/comprobantesFiscalesElectronicos.aspx) |
| R3-CREDITO-SIN-COMPRADOR | ERROR | B01/E31 requiere RNC/cédula válido del comprador | DGII Guía Informativa NCF |
| R3-CONSUMO-CON-RNC | ADVERTENCIA | B02/E32 emitida a un comprador con RNC (no puede usarla como crédito de ITBIS ni gasto de ISR) | DGII Guía Informativa NCF |
| R3-EXENTO-CON-ITBIS | ADVERTENCIA | B14/E44 o B16/E46 con ITBIS cobrado (normalmente exento / 0%) | DGII [Guía Comprobantes Especiales](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/3-Guia-Comprobantes-Fiscales-Especiales-NG-05-19.pdf) |
| R3-NOTA-SIN-REFERENCIA | ADVERTENCIA | B03/B04/E33/E34 sin `ncf_modificado` | DGII Guía Informativa NCF |
| R4-NCF-DUPLICADO | ERROR | mismo emisor + NCF en facturas con distinta fecha o comprador | DGII Guía Informativa NCF (cada NCF es único) |
| R4-LINEA-REPETIDA | ADVERTENCIA | línea idéntica repetida dentro de una factura (posible doble registro) | heurística de esta herramienta |
| R4-SECUENCIA-SALTO | INFO | saltos en la secuencia por emisor + tipo (pueden ser anulados; se reportan en el formato 608) | heurística de esta herramienta |
| R4-SECUENCIA-VENCIDA | ERROR | `fecha` posterior a `vencimiento_secuencia` | DGII Tipos de comprobantes |
| R5-TASA-INVALIDA | ERROR | la tasa es una de `rules/itbis_rates.yaml`: 18%, 16%, 0%, exento | DGII [ITBIS](https://dgii.gov.do/cicloContribuyente/obligacionesTributarias/principalesImpuestos/Paginas/Itbis.aspx), [Guía #7 ITBIS](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/itbis/Documents/1-Guia%207%20-%20(ITBIS).pdf) |
| R5-NUMERO-INVALIDO | ERROR | cantidad, precio, ITBIS o total vacío o ilegible | validación de datos |
| R5-ITBIS-LINEA | ERROR | `cantidad × precio_unitario × tasa` vs ITBIS cobrado, tolerancia RD$0.01 por línea | DGII ITBIS, Guía #7 (la tolerancia es una decisión de esta herramienta) |
| R5-TOTAL-FACTURA | ERROR | Σ total vs Σ base + Σ ITBIS, tolerancia RD$0.05 por factura | aritmética (la tolerancia es una decisión de esta herramienta) |
| R5-TASA-REDUCIDA | INFO | toda línea al 16%: confirmar que el producto está en la lista de tasa reducida | DGII Guía #7; ver la nota de la Ley 30-26 en `rules/itbis_rates.yaml` |
| R6-FECHA-INVALIDA | ERROR | fecha vacía o inexistente | validación de datos |
| R6-FECHA-FUTURA | ERROR | fecha posterior a hoy | validación de datos |

**Sobre python-stdnum y los tipos de NCF.** El verificador también usa `stdnum.do.ncf`. Si
stdnum rechaza un tipo que aparece en `rules/ncf_types.yaml` (stdnum puede atrasarse respecto
a la DGII), prevalece el YAML y el resumen lo indica. Con python-stdnum 2.2 ambas listas coinciden.

**Ley 30-26.** La Ley 30-26 (junio 2026) agregó nuevas exenciones de ITBIS para bienes seleccionados, con efecto inmediato. Los resúmenes publicados de la reforma no indican cambios en las tasas de 18% / 16%, pero esto no se ha confirmado contra el texto de la ley. Esta herramienta no revisa listas de productos: si un producto podría estar exento ahora, o si tienes dudas sobre una tasa, confírmalo con tu contador. Fuente: [RSM, Ley 30-26](https://www.rsm.global/dominicanrepublic/en/news/dominican-republic-tax-reform-law-30-26).

### Limitaciones

- **No consulta a la DGII en línea.** No puede saber si un RNC está activo, si un NCF fue
  autorizado para ese emisor ni si un e-CF fue aceptado; solo si están bien formados.
- No sabe qué productos tienen tasa de 16% o están exentos; marca las líneas al 16% para revisión.
- La advertencia de B02/E32 solo aplica a RNC de 9 dígitos, no a cédulas.
- Solo se aceptan las series `B` y `E` (la DGII permite otras letras de serie e-CF en algunos casos).
- Se asume que los precios no incluyen ITBIS. No se revisan ISC, propina legal ni otros impuestos.
- La detección de duplicados se basa en diferencias de fecha/comprador; una factura pegada dos
  veces con líneas idénticas se reporta como líneas repetidas (ADVERTENCIA), no como NCF duplicado.

### Desarrollo

```bash
python -m pip install -r requirements.txt pytest
python -m pytest
python scripts/make_template.py   # regenera la plantilla de Excel
```

### Aviso legal

Este proyecto no es asesoría fiscal y no está afiliado a la DGII. Las tasas y reglas cambian;
confirma siempre con la DGII y revisa los resultados con tu contador antes de presentar
cualquier declaración.

Licencia: [MIT](LICENSE).
