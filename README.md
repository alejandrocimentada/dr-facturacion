**Español** · [English](README.en.md)

![DR Facturación: revisa tus facturas antes de declarar](docs/banner.png)

<p align="center"><strong>Revisa RNC, NCF/e-NCF e ITBIS de tus facturas en segundos, y recibe un reporte en Excel con cada problema explicado en español.</strong></p>

## ¿Qué es esto?

Le das tus facturas en un archivo de Excel. Te dice cuáles tienen errores y cómo corregirlos.
Nunca cambia tu archivo: solo lo lee y crea un reporte nuevo aparte.

## ¿Para quién es?

Para dueños de negocios, contadores y profesionales independientes en República Dominicana
que quieren encontrar errores en sus facturas antes de declarar o de enviárselas al contador.

Si eres pequeño o micro contribuyente, tienes hasta el **15 de noviembre de 2026** para empezar
a emitir facturas electrónicas (e-CF).

> [!IMPORTANT]
> **No cambia tu archivo.** Solo lo lee y crea un reporte nuevo.
> **Tus datos no salen de tu computadora.** No se conecta a la DGII ni a ningún otro lugar.
> **No es asesoría fiscal.** Revisa los resultados con tu contador antes de declarar.

## Cómo usarlo (la forma más fácil)

1. **Prepara tus facturas en Excel.**
   [Descarga la plantilla](https://github.com/alejandrocimentada/dr-facturacion/raw/main/templates/plantilla_facturas.xlsx)
   y llénala, con una fila por cada producto o servicio de cada factura.
   También puedes usar tu propio Excel si tiene columnas parecidas (fecha, RNC, NCF, cantidad,
   precio, ITBIS, total).

2. **Pídele a tu asistente de IA que lo revise.** Si usas Claude Code u otro agente de IA
   (un asistente que puede trabajar con archivos en tu computadora), copia y pega este mensaje:

   ```text
   Instala el skill dr-facturacion desde https://github.com/alejandrocimentada/dr-facturacion
   y luego revisa mi archivo facturas.xlsx
   ```

   El agente lo instala todo por ti. Cambia `facturas.xlsx` por el nombre de tu archivo.

   <details>
   <summary><b>¿No tienes un agente de IA? Así empiezas (5 minutos)</b></summary>

   1. **Qué es un agente de IA:** un asistente, como Claude Code, que puede trabajar con archivos
      en tu computadora: leerlos, ejecutar programas y explicarte los resultados.
   2. **Abre la terminal** (la ventana donde se escriben comandos):
      - Windows: presiona la tecla Windows, escribe "PowerShell" y ábrelo.
      - Mac: presiona Cmd + Espacio, escribe "Terminal" y ábrela.
   3. **Instala Claude Code** siguiendo la
      [guía oficial de inicio rápido](https://code.claude.com/docs/es/quickstart) (pasos 1 y 2).
      Necesitarás una cuenta de Claude; la guía explica qué planes incluyen Claude Code.
   4. **Ve a la carpeta de tus facturas** en la terminal. Ejemplo:
      - Windows: `cd "$HOME\Documents\Facturas"`
      - Mac: `cd ~/Documents/Facturas`
   5. **Escribe `claude`** y presiona Enter para abrir el agente. La primera vez te pedirá
      iniciar sesión con tu cuenta de Claude. Luego pega el mensaje del paso 2.

   </details>

3. **Abre el reporte y corrige.** El agente crea un archivo nuevo llamado
   `facturas_reporte.xlsx` junto a tu archivo. Ábrelo y corrige las filas marcadas en
   rojo (errores). Las filas en naranja son advertencias: revísalas también.

El reporte tiene tres pestañas:

**Resumen**: cuántas facturas se revisaron, cuántos problemas hay y si el ITBIS cuadra.

![Pestaña Resumen](docs/reporte-resumen.png)

**Problemas**: un problema por fila, con qué está mal y qué hacer. Los más graves van primero.

![Pestaña Problemas](docs/reporte-problemas.png)

**Facturas**: tus datos tal como estaban, con cada fila en color según su estado (rojo, naranja o verde).

![Pestaña Facturas](docs/reporte-facturas.png)

## ¿Qué significa cada error?

| Error | Qué significa | Qué hacer |
|---|---|---|
| El RNC o la cédula del emisor no es válido | Probablemente hay un número mal escrito, o falta. | Copia el RNC exacto desde una factura o documento oficial. |
| El RNC o la cédula del comprador no es válido | Probablemente hay un número mal escrito. | Pídele el RNC o la cédula correcta al cliente. |
| Cédula sin ceros al inicio | Excel borró los ceros del principio. La herramienta los completó por ti. | Nada urgente. Para evitarlo, guarda esa columna como texto. |
| El NCF tiene un formato incorrecto | Al número de comprobante le faltan o le sobran caracteres. | Cópialo de nuevo desde la factura original. |
| Tipo de comprobante desconocido | Las letras y números del inicio del NCF (como B01) no existen. | Revisa que el NCF esté bien escrito. |
| Crédito fiscal sin RNC del comprador | Una factura B01/E31 siempre debe llevar el RNC de quien compra. | Agrega el RNC del cliente, o usa una factura de consumo (B02). |
| Factura de consumo a una empresa | Diste una factura B02/E32 a alguien que tiene RNC. | Confirma con el cliente: tal vez necesitaba crédito fiscal (B01). |
| ITBIS en una factura que no debería llevarlo | Facturas de régimen especial o exportación casi nunca llevan ITBIS. | Revisa con tu contador si ese ITBIS va o no. |
| Nota de crédito o débito sin factura original | La nota no dice a qué factura corrige. | Escribe el NCF de la factura original en la columna `ncf_modificado`. |
| NCF repetido | El mismo número de comprobante aparece en dos facturas distintas. | Revisa cuál es el correcto. Cada factura debe tener su propio NCF. |
| Línea repetida | La misma línea aparece dos veces en una factura. | Bórrala si se copió por error. |
| Salto en la numeración | Faltan números entre un NCF y el siguiente. | Puede que sean facturas anuladas. Si es así, deben ir en el reporte de anulados de la DGII (formato 608). |
| Secuencia vencida | La factura se hizo después de la fecha de vencimiento de esos NCF. | Pide una secuencia nueva a la DGII y habla con tu contador. |
| Tasa de ITBIS no válida | La tasa no es 18%, 16%, 0% ni "exento". | Corrige la tasa. Lo normal es 18%. |
| Número vacío o ilegible | Falta un monto o tiene letras donde van números. | Llena la cantidad, el precio, el ITBIS y el total. |
| El ITBIS no cuadra | El ITBIS no da cantidad × precio × tasa. | Recalcula el ITBIS de esa línea. |
| El total no cuadra | El total de la factura no da el monto sin ITBIS + el ITBIS. | Recalcula el total. |
| Línea con tasa del 16% | Solo algunos productos llevan 16%. Es un aviso, no un error. | Confirma que ese producto de verdad lleva 16%. |
| Fecha vacía o inválida | Falta la fecha, o no existe (por ejemplo, 31 de febrero). | Escribe la fecha correcta (DD/MM/AAAA). |
| Fecha en el futuro | La fecha es posterior a hoy. | Corrige el año, el mes o el día. |

## Glosario

- **RNC**: Registro Nacional de Contribuyentes. El número de 9 dígitos con el que la DGII identifica a una empresa.
- **Cédula**: el número de identidad de 11 dígitos de una persona. Sirve como RNC para personas físicas.
- **NCF**: Número de Comprobante Fiscal. El número oficial de cada factura en papel, como `B0100000001`.
- **e-NCF / e-CF**: la versión electrónica. El e-CF es la factura electrónica; el e-NCF es su número, como `E310000000001`.
- **ITBIS**: el impuesto que se cobra sobre la venta de bienes y servicios. Normalmente es 18%.
- **Crédito fiscal**: el ITBIS que una empresa pagó al comprar y que puede descontar del que debe pagar.
- **B01 vs B02**: B01 es factura de crédito fiscal (para empresas, lleva el RNC del comprador). B02 es factura de consumo (para personas, sin RNC).

## Preguntas frecuentes

**¿Mis datos se envían a algún lugar?**
No. Todo se revisa en tu computadora. No se conecta a la DGII ni a internet.

**¿Reemplaza a mi contador?**
No. Te ayuda a encontrar errores antes de enviarle las facturas. La decisión final es de tu contador.

**¿Puede leer facturas en PDF?**
Todavía no. Por ahora solo lee Excel (.xlsx) y CSV (un formato de tabla que Excel puede guardar).

**¿Revisa si un RNC está activo?**
No. Solo revisa que el número esté bien formado. Para saber si está activo, búscalo en la página de la DGII.

**¿Necesito saber programar?**
No, si usas un agente de IA como Claude Code. El agente se encarga de instalarlo y usarlo.

---

## Para usuarios técnicos

**dr-facturacion** es una skill para agentes de IA (Claude Code, Codex y otros agentes
compatibles con `SKILL.md`) que también funciona como comando independiente. El plazo de
e-CF del 15 de noviembre de 2026 viene del Aviso 06-26 de la DGII
([resumen](https://siemprealdia.co/republica-dominicana/impuestos/dgii-prorrogo-el-e-cf-para-pequenos-micros-y-no-clasificados/)).

### Instalación manual

Requisitos: Python 3.10 o superior y `pip`.

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

Una vez instalada, pídeselo al agente en una conversación:

> Usa dr-facturacion para revisar mis facturas de septiembre: facturas_septiembre.xlsx

El agente ejecuta la revisión, te explica los problemas principales agrupados por severidad
con los números de fila, y te indica dónde está el reporte.

### Uso como comando (sin agente)

```bash
git clone https://github.com/alejandrocimentada/dr-facturacion.git
cd dr-facturacion
pip install -r requirements.txt
python scripts/check_invoices.py facturas.xlsx
```

El reporte se guarda junto al archivo como `facturas_reporte.xlsx`. Otras opciones:
`--sheet "Septiembre"` (hoja de Excel), `--output revision.xlsx` (ruta del reporte) y
`--map ncf="No. Comprobante"` (asignar una columna que no se reconoce).

### Columnas y mapeo (`--map`)

Usa [`templates/plantilla_facturas.xlsx`](templates/plantilla_facturas.xlsx): una fila por
cada línea de factura. Los encabezados pueden estar en español o inglés (se ignoran
mayúsculas, acentos y signos), y CSV o Excel funcionan igual. Si una columna tiene otro
nombre, asígnala con `--map campo="Nombre de tu columna"`.

| Columna | Obligatoria | Descripción |
|---|---|---|
| `fecha` | sí | fecha de emisión (DD/MM/AAAA) |
| `rnc_emisor` | sí | RNC o cédula de quien emite |
| `rnc_comprador` | sí (puede ir vacía en consumo) | RNC o cédula del comprador |
| `ncf` | sí | NCF o e-NCF |
| `descripcion` | no | producto o servicio |
| `cantidad` | sí | cantidad |
| `precio_unitario` | sí | precio unitario **sin** ITBIS |
| `tasa_itbis` | sí | `18`, `16`, `0` o `exento` |
| `itbis` | sí | ITBIS cobrado en la línea |
| `total` | sí | total de la línea (base + ITBIS) |
| `vencimiento_secuencia` | no | vencimiento de la secuencia de NCF |
| `ncf_modificado` | no | NCF original, en notas de débito/crédito |

Las capturas del reporte de arriba se generaron a partir de
[`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv).

### ¿Qué revisa?

| Revisión | Qué detecta | Severidad |
|---|---|---|
| RNC / cédula | RNC (9 dígitos) o cédula (11 dígitos) del emisor o comprador vacío, mal formado o con dígito verificador inválido; cédulas que perdieron los ceros iniciales en Excel | ERROR · INFO |
| Estructura NCF / e-NCF | NCF que no sigue `B` + tipo + 8 dígitos, e-NCF que no sigue `E` + tipo + 10 dígitos, tipos de comprobante que no existen, formato anterior a 2018 | ERROR |
| Tipo de NCF vs comprador | crédito fiscal (B01/E31) sin RNC del comprador; consumo (B02/E32) emitido a una empresa con RNC; régimen especial o exportación (B14/E44, B16/E46) con ITBIS; notas de débito/crédito sin NCF modificado | ERROR · ADVERTENCIA |
| Secuencias | NCF duplicados, líneas repetidas, secuencias vencidas, saltos en la numeración (posibles anulados para el formato 608) | ERROR · ADVERTENCIA · INFO |
| Cálculo de ITBIS | tasas que no son 18%, 16%, 0% o exento; ITBIS de la línea distinto de cantidad × precio × tasa; montos ilegibles; líneas al 16% para confirmar | ERROR · INFO |
| Totales | total de la factura distinto de base + ITBIS | ERROR |
| Fechas | fechas vacías, inexistentes (p. ej. 31/02) o en el futuro | ERROR |

El detalle de cada una de las 20 reglas está en [`scripts/check_invoices.py`](scripts/check_invoices.py) (diccionario `RULES`).

### ¿Hasta qué punto está verificado?

- **Reglas verificadas contra fuentes de la DGII el 2026-10-07** (campo `last_verified` en
  [`rules/itbis_rates.yaml`](rules/itbis_rates.yaml) y [`rules/ncf_types.yaml`](rules/ncf_types.yaml)):
  - [DGII: ITBIS](https://dgii.gov.do/cicloContribuyente/obligacionesTributarias/principalesImpuestos/Paginas/Itbis.aspx)
  - [DGII: Guía #7 ITBIS (PDF)](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/itbis/Documents/1-Guia%207%20-%20%28ITBIS%29.pdf)
  - [DGII: Tipos de comprobantes fiscales](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/tiposComprobantes.aspx)
  - [DGII: Guía Informativa sobre Comprobantes Fiscales (PDF)](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/2-Guia-Informativa-NCF.pdf)
  - [DGII: Guía de Comprobantes Fiscales Especiales (PDF)](https://dgii.gov.do/publicacionesOficiales/bibliotecaVirtual/contribuyentes/facturacion/Documents/Comprobantes%20Fiscales/3-Guia-Comprobantes-Fiscales-Especiales-NG-05-19.pdf)
  - [DGII: Comprobantes Fiscales Electrónicos (e-CF)](https://dgii.gov.do/cicloContribuyente/facturacion/comprobantesFiscales/Paginas/comprobantesFiscalesElectronicos.aspx)
  - [DGII: Aviso 24-19, estructura del e-NCF (PDF)](https://dgii.gov.do/publicacionesOficiales/avisosInformativos/Documents/2019/24-19.pdf)
  - Dígitos verificadores: [python-stdnum `stdnum.do.rnc`](https://arthurdejong.org/python-stdnum/doc/2.2/stdnum.do.rnc.html) y [`stdnum.do.cedula`](https://arthurdejong.org/python-stdnum/doc/2.2/stdnum.do.cedula.html)
- **140 pruebas automáticas**, con casos válidos e inválidos para cada regla.
- **Probado solo con el archivo de ejemplo** ([`examples/facturas_ejemplo.csv`](examples/facturas_ejemplo.csv)),
  todavía no con facturas reales de empresas.
- Las reglas de **saltos de secuencia y líneas repetidas**, y las **tolerancias** de redondeo
  (RD$0.01 por línea, RD$0.05 por factura), son heurísticas de esta herramienta, no reglas de la DGII.
- **Ley 30-26:** La Ley 30-26 (junio 2026) agregó nuevas exenciones de ITBIS para bienes seleccionados, con efecto inmediato. Los resúmenes publicados de la reforma no indican cambios en las tasas de 18% / 16%, pero esto no se ha confirmado contra el texto de la ley. Esta herramienta no revisa listas de productos: si un producto podría estar exento ahora, o si tienes dudas sobre una tasa, confírmalo con tu contador.
  ([Fuente: RSM, Ley 30-26](https://www.rsm.global/dominicanrepublic/en/news/dominican-republic-tax-reform-law-30-26))

### ¿Por qué es confiable?

1. **Nunca toca el archivo de entrada.** Lo lee en memoria y escribe un reporte aparte; las
   pruebas verifican con un hash que el archivo queda idéntico (CSV y Excel).
2. **Los dígitos verificadores se validan con [python-stdnum](https://arthurdejong.org/python-stdnum/)**,
   una biblioteca abierta y mantenida, no con fórmulas propias.
3. **Las reglas viven en archivos YAML editables** ([`rules/`](rules/)), con su fecha de
   verificación y sus fuentes. Si las tasas tienen más de 90 días sin verificarse, la herramienta lo avisa.
4. **Cada problema explica qué está mal y qué hacer**, en español sencillo, con el número de fila y el NCF.

### Personalizar las reglas

Las reglas están en dos archivos YAML que puedes editar sin tocar el código:

- [`rules/itbis_rates.yaml`](rules/itbis_rates.yaml): tasas de ITBIS aceptadas (`general`,
  `reducida`, `exportacion`, `exento`), la nota sobre la Ley 30-26 (`note` / `nota_es`) y sus fuentes.
- [`rules/ncf_types.yaml`](rules/ncf_types.yaml): series (`B`, `E`) con su longitud, y los tipos de
  comprobante válidos con sus marcas (`requires_buyer_id`, `consumer`, `usually_exempt`, `modifies`).

Cuando cambies algo, **actualiza `last_verified`** con la fecha de hoy y **agrega la fuente**
en `sources`, para que quien use la herramienta sepa de dónde sale cada regla.

### Desarrollo y validación

```bash
pip install -r requirements.txt pytest
python -m pytest                                          # 140 pruebas
python scripts/check_invoices.py examples/facturas_ejemplo.csv
python scripts/make_template.py                           # regenera la plantilla de Excel
```

El archivo de ejemplo activa las 20 reglas al menos una vez. El banner se genera desde
[`docs/banner.html`](docs/banner.html).

### Límites conocidos

- **No consulta a la DGII en línea**: no puede confirmar que un RNC esté activo ni que un NCF
  haya sido autorizado, solo que estén bien formados.
- **No revisa listas de productos**: no sabe qué productos tienen tasa de 16% o están exentos.
- **Todavía no lee PDF**: solo CSV y Excel (.xlsx).
- **No cubre retenciones** de ITBIS ni de ISR (ni ISC o propina legal).
- **Las reglas pueden cambiar**: revisa la fecha `last_verified` en [`rules/`](rules/) antes de confiar en los resultados.

## Licencia

[MIT](LICENSE) © 2026 Alejandro Cimentada.
