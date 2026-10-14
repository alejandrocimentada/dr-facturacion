// DR Facturación web app: UI only. The checking happens in worker.js (Pyodide), in this browser.
(() => {
  "use strict";

  const I18N = {
    es: {
      skip: "Saltar al contenido",
      lang: "Idioma",
      github: "GitHub",
      eyebrow: "REVISOR DE FACTURAS · RNC · NCF · ITBIS",
      title: "Revisa tus facturas antes de declarar.",
      lead: "Suelta tu Excel o CSV y en segundos verás qué RNC, NCF o ITBIS tienen errores, y cómo corregirlos.",
      dropTitle: "Arrastra tu archivo aquí",
      dropTitleTouch: "Elige tu archivo de facturas",
      dropOr: "o",
      dropButton: "Elegir archivo",
      dropHint: ".xlsx o .csv · una fila por cada línea de factura",
      dropActive: "Suelta para revisar",
      privacy: "Tu archivo nunca sale de tu navegador.",
      privacyDetail: "La revisión corre en esta página con Python (Pyodide). No se sube a ningún servidor ni se conecta a la DGII.",
      template: "Descargar plantilla (.xlsx)",
      example: "Probar con un archivo de ejemplo",
      illustrationAlt: "Un niño revisa una factura con una lupa junto a un gecko verde sentado sobre una pila de papeles. La factura tiene tres marcas verdes y una roja.",
      warmLoading: "Preparando el revisor en segundo plano · {n}/4",
      warmReady: "Revisor listo · todo corre en tu navegador",
      warmSaveData: "El revisor se descargará cuando elijas un archivo (ahorro de datos activado)",
      loadingEyebrow: "UN MOMENTO",
      loadingTitle: "Preparando la revisión",
      stepPython: "Descargando Python (Pyodide)",
      stepPandas: "Cargando pandas y PyYAML",
      stepLibs: "Instalando openpyxl y python-stdnum",
      stepRules: "Cargando las reglas de la DGII",
      stepCheck: "Revisando tu archivo",
      loadingNote: "La primera vez el navegador descarga unos 16 MB (Python y pandas). Después quedan guardados y la revisión es mucho más rápida.",
      errorEyebrow: "NO SE PUDO REVISAR",
      tryAgain: "Elegir otro archivo",
      retry: "Reintentar",
      errColumnsTitle: "No reconocimos algunas columnas",
      errColumnsBody: "Faltan estas columnas:",
      errColumnsFound: "Encabezados en tu archivo:",
      errColumnsFix: "Cambia los encabezados para que se parezcan a los de la plantilla (por ejemplo «NCF», «RNC Emisor», «Tasa ITBIS»), o copia tus datos en la plantilla.",
      errReadTitle: "No pudimos leer el archivo",
      errReadBody: "Asegúrate de que sea un Excel (.xlsx) o CSV válido y que no esté protegido con contraseña.",
      errDetail: "Detalle técnico:",
      errTypeTitle: "Ese tipo de archivo no se puede revisar",
      errTypeBody: "Usa un archivo .xlsx o .csv. Si es un Excel antiguo (.xls), ábrelo y guárdalo como .xlsx.",
      errEmptyTitle: "El archivo no tiene facturas",
      errEmptyBody: "No encontramos filas con datos debajo de los encabezados.",
      errBootTitle: "No se pudo cargar el revisor",
      errBootBody: "La primera vez hace falta internet para descargar Python en el navegador. Revisa tu conexión y vuelve a intentarlo. Tu archivo no se envió a ningún lado.",
      resultsEyebrow: "RESULTADO",
      resultsTitle: "Revisión de {file}",
      meta: "Revisado el {date} · tasas ITBIS verificadas el {verified} (hace {age} días)",
      download: "Descargar reporte Excel",
      again: "Revisar otro archivo",
      stale: "Atención: las tasas de ITBIS no se verifican desde hace {age} días; confírmalas en dgii.gov.do.",
      cardInvoices: "Facturas revisadas",
      cardInvoicesSub: "{lines} líneas · {ok} OK",
      cardErrors: "Errores",
      cardErrorsSub: "corregir antes de declarar",
      cardWarnings: "Advertencias",
      cardWarningsSub: "revisar con atención",
      cardInfo: "Info",
      cardInfoSub: "solo confirmar",
      cardItbis: "Diferencia ITBIS",
      cardItbisSub: "cobrado {charged} · recalculado {expected}",
      todo: "QUÉ HACER AHORA",
      todoNone: "No encontramos problemas en este archivo.",
      problems: "PROBLEMAS",
      invoices: "FACTURAS",
      all: "Todos",
      allRows: "Todas",
      severityFilter: "Filtrar por severidad",
      statusFilter: "Filtrar por estado",
      tableLabel: "Tabla de facturas",
      searchLabel: "Buscar problemas",
      search: "Buscar fila, NCF o texto",
      noMatch: "Ningún problema coincide con el filtro.",
      more: "Mostrar {n} más",
      hSeverity: "Severidad", hRow: "Fila", hNcf: "NCF", hProblem: "Problema", hDo: "Qué hacer", hStatus: "Estado",
      detailEs: "",
      row: "Fila", rows: "Filas", andMore: "y {n} más",
      goRow: "Ver fila {n} en la tabla de facturas",
      rowsLimited: "Se muestran las primeras {shown} de {total} filas. El reporte Excel tiene todas.",
      announce: "Revisión lista: {e} errores, {w} advertencias, {i} avisos informativos.",
      sev: { ERROR: "ERROR", ADVERTENCIA: "ADVERTENCIA", INFO: "INFO", OK: "OK" },
      sevPlural: { ERROR: "Errores", ADVERTENCIA: "Advertencias", INFO: "Info", OK: "OK" },
      disclaimer: "No es asesoría fiscal. Revisa estos resultados con tu contador antes de presentar cualquier declaración.",
      foot: "Código abierto (MIT) · Las mismas reglas que la versión de terminal",
    },
    en: {
      skip: "Skip to content",
      lang: "Language",
      github: "GitHub",
      eyebrow: "DOMINICAN INVOICE CHECKER · RNC · NCF · ITBIS",
      title: "Check your invoices before you file.",
      lead: "Drop your Excel or CSV file and see in seconds which RNC, NCF or ITBIS values are wrong, and how to fix them.",
      dropTitle: "Drag your file here",
      dropTitleTouch: "Choose your invoice file",
      dropOr: "or",
      dropButton: "Choose a file",
      dropHint: ".xlsx or .csv · one row per invoice line",
      dropActive: "Drop to check",
      privacy: "Your file never leaves your browser.",
      privacyDetail: "The check runs on this page with Python (Pyodide). Nothing is uploaded to a server, and it never connects to the DGII.",
      template: "Download the template (.xlsx)",
      example: "Try it with an example file",
      illustrationAlt: "A kid checks an invoice with a magnifying glass next to a green gecko sitting on a stack of papers. The invoice has three green check marks and one red cross.",
      warmLoading: "Getting the checker ready in the background · {n}/4",
      warmReady: "Checker ready · everything runs in your browser",
      warmSaveData: "The checker will download when you choose a file (data saver is on)",
      loadingEyebrow: "ONE MOMENT",
      loadingTitle: "Getting the check ready",
      stepPython: "Downloading Python (Pyodide)",
      stepPandas: "Loading pandas and PyYAML",
      stepLibs: "Installing openpyxl and python-stdnum",
      stepRules: "Loading the DGII rules",
      stepCheck: "Checking your file",
      loadingNote: "The first time, your browser downloads about 16 MB (Python and pandas). After that it is cached and checks are much faster.",
      errorEyebrow: "COULDN'T CHECK THE FILE",
      tryAgain: "Choose another file",
      retry: "Try again",
      errColumnsTitle: "Some columns weren't recognized",
      errColumnsBody: "These columns are missing:",
      errColumnsFound: "Headers in your file:",
      errColumnsFix: "Rename your headers to look like the template's (for example “NCF”, “RNC Emisor”, “Tasa ITBIS”), or paste your data into the template.",
      errReadTitle: "We couldn't read the file",
      errReadBody: "Make sure it's a valid Excel (.xlsx) or CSV file and that it isn't password-protected.",
      errDetail: "Technical detail:",
      errTypeTitle: "That file type can't be checked",
      errTypeBody: "Use an .xlsx or .csv file. If it's an old Excel file (.xls), open it and save it as .xlsx.",
      errEmptyTitle: "The file has no invoices",
      errEmptyBody: "We didn't find any rows with data below the headers.",
      errBootTitle: "The checker couldn't load",
      errBootBody: "The first time, an internet connection is needed to download Python into your browser. Check your connection and try again. Your file wasn't sent anywhere.",
      resultsEyebrow: "RESULT",
      resultsTitle: "Check of {file}",
      meta: "Checked on {date} · ITBIS rates verified on {verified} ({age} days ago)",
      download: "Download Excel report",
      again: "Check another file",
      stale: "Warning: the ITBIS rates haven't been verified for {age} days; confirm them at dgii.gov.do.",
      cardInvoices: "Invoices checked",
      cardInvoicesSub: "{lines} lines · {ok} OK",
      cardErrors: "Errors",
      cardErrorsSub: "fix before filing",
      cardWarnings: "Warnings",
      cardWarningsSub: "review carefully",
      cardInfo: "Info",
      cardInfoSub: "just confirm",
      cardItbis: "ITBIS difference",
      cardItbisSub: "charged {charged} · recalculated {expected}",
      todo: "WHAT TO DO NOW",
      todoNone: "We didn't find any problems in this file.",
      problems: "PROBLEMS",
      invoices: "INVOICES",
      all: "All",
      allRows: "All",
      severityFilter: "Filter by severity",
      statusFilter: "Filter by status",
      tableLabel: "Invoices table",
      searchLabel: "Search problems",
      search: "Search row, NCF or text",
      noMatch: "No problems match the filter.",
      more: "Show {n} more",
      hSeverity: "Severity", hRow: "Row", hNcf: "NCF", hProblem: "Problem", hDo: "What to do", hStatus: "Status",
      detailEs: "Detail (in Spanish)",
      row: "Row", rows: "Rows", andMore: "and {n} more",
      goRow: "Show row {n} in the invoices table",
      rowsLimited: "Showing the first {shown} of {total} rows. The Excel report has all of them.",
      announce: "Check finished: {e} errors, {w} warnings, {i} info notes.",
      sev: { ERROR: "ERROR", ADVERTENCIA: "WARNING", INFO: "INFO", OK: "OK" },
      sevPlural: { ERROR: "Errors", ADVERTENCIA: "Warnings", INFO: "Info", OK: "OK" },
      disclaimer: "This is not tax advice. Review these results with your accountant before filing any return.",
      foot: "Open source (MIT) · Same rules as the command-line version",
    },
  };

  // English title + action per rule. Spanish comes from the checker itself (RULE_HELP).
  const RULES_EN = {
    "R1-RNC-EMISOR": ["The issuer's RNC or cédula isn't valid", "Copy the exact RNC from an invoice or official document."],
    "R1-RNC-COMPRADOR": ["The buyer's RNC or cédula isn't valid", "Ask the customer for the correct RNC or cédula."],
    "R1-CEDULA-CEROS": ["Cédula missing its leading zeros", "Nothing urgent. To avoid it, store that column as text."],
    "R2-NCF-ESTRUCTURA": ["The NCF has the wrong format", "Copy the NCF again from the original invoice."],
    "R2-NCF-TIPO": ["Unknown receipt type", "Check that the NCF is written correctly."],
    "R3-CREDITO-SIN-COMPRADOR": ["Fiscal-credit invoice without the buyer's RNC", "Add the customer's RNC, or use a consumer invoice (B02)."],
    "R3-CONSUMO-CON-RNC": ["Consumer invoice issued to a company", "Check with the customer: they may have needed a fiscal-credit invoice (B01)."],
    "R3-EXENTO-CON-ITBIS": ["ITBIS on an invoice that usually has none", "Check with your accountant whether that ITBIS applies."],
    "R3-NOTA-SIN-REFERENCIA": ["Credit or debit note without the original invoice", "Write the original invoice's NCF in the ncf_modificado column."],
    "R4-NCF-DUPLICADO": ["Repeated NCF", "Check which one is right. Each invoice needs its own NCF."],
    "R4-LINEA-REPETIDA": ["Repeated line", "Delete it if it was copied by mistake."],
    "R4-SECUENCIA-SALTO": ["Gap in the numbering", "If they are voided invoices, they must be reported in DGII form 608."],
    "R4-SECUENCIA-VENCIDA": ["Expired sequence", "Request a new sequence from the DGII and talk to your accountant."],
    "R5-TASA-INVALIDA": ["Invalid ITBIS rate", "Fix the rate: 18%, 16%, 0% or exento. The usual rate is 18%."],
    "R5-NUMERO-INVALIDO": ["Empty or unreadable number", "Fill in quantity, price, ITBIS and total with numbers."],
    "R5-ITBIS-LINEA": ["The ITBIS doesn't add up", "Recalculate the line's ITBIS: quantity × price × rate."],
    "R5-TOTAL-FACTURA": ["The total doesn't add up", "Recalculate the total: amount before ITBIS + ITBIS."],
    "R5-TASA-REDUCIDA": ["Line at the 16% rate", "Confirm that this product really takes 16%."],
    "R6-FECHA-INVALIDA": ["Empty or invalid date", "Write the correct date (DD/MM/YYYY)."],
    "R6-FECHA-FUTURA": ["Date in the future", "Fix the year, month or day."],
  };
  const SEVERITIES = ["ERROR", "ADVERTENCIA", "INFO"];
  const STEPS = ["python", "pandas", "libs", "rules", "check"];
  const PAGE = 150;           // problems rendered per "show more"
  const MAX_ROWS = 1000;      // invoice rows rendered in the table
  const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const $ = (sel) => document.querySelector(sel);
  const el = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else if (k === "style") node.style.cssText = v;
      else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v === true ? "" : v);
    }
    for (const child of children.flat()) {
      if (child == null) continue;
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    }
    return node;
  };
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");

  const storage = {
    get(key) { try { return localStorage.getItem(key); } catch { return null; } },
    set(key, value) { try { localStorage.setItem(key, value); } catch { /* private mode */ } },
  };

  let lang = storage.get("drf-lang") === "en" ? "en" : "es";
  const lawNote = { es: null, en: null };
  const state = { result: null, reportUrl: null, sev: "ALL", query: "", rowFilter: "ALL", shown: PAGE, lastFile: null };
  let screen = "welcome";
  let ready = false;
  let bootStarted = false;
  let warmStep = 0;
  let runId = 0;
  let loadingTimer = null;

  const touchOnly = window.matchMedia("(hover: none)").matches;
  if (touchOnly) document.getElementById("drop-title").dataset.i18n = "dropTitleTouch";

  function t(key, vars = {}) {
    const text = I18N[lang][key] ?? I18N.es[key] ?? key;
    return typeof text === "string" ? text.replace(/\{(\w+)\}/g, (_, k) => vars[k] ?? "") : text;
  }
  const fmtInt = (n) => Number(n).toLocaleString("en-US");
  const money = (n) => (n < -0.004 ? "-" : "") + "RD$" + Math.abs(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fold = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const ruleTitle = (rule, fallback) => (lang === "en" ? RULES_EN[rule]?.[0] : null) ?? fallback;
  const ruleAction = (rule, fallback) => (lang === "en" ? RULES_EN[rule]?.[1] : null) ?? fallback;

  // ------------------------------------------------------------------ language
  function applyLanguage() {
    document.documentElement.lang = lang;
    document.querySelectorAll("[data-i18n]").forEach((n) => { n.textContent = t(n.dataset.i18n); });
    document.querySelectorAll("[data-i18n-placeholder]").forEach((n) => { n.placeholder = t(n.dataset.i18nPlaceholder); });
    document.querySelectorAll("[data-i18n-alt]").forEach((n) => { n.alt = t(n.dataset.i18nAlt); });
    document.querySelectorAll("[data-i18n-aria]").forEach((n) => { n.setAttribute("aria-label", t(n.dataset.i18nAria)); });
    document.querySelectorAll(".lang-btn").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.lang === lang)));
    if (lawNote[lang]) $("#law-note").textContent = lawNote[lang];
    updateWarm();
    if (state.result && screen === "results") renderResults({ animate: false });
  }
  document.querySelectorAll(".lang-btn").forEach((b) => b.addEventListener("click", () => {
    if (b.dataset.lang === lang) return;
    lang = b.dataset.lang;
    storage.set("drf-lang", lang);
    applyLanguage();
  }));

  // The Ley 30-26 note comes from the same rules file the checker uses.
  function foldedScalar(yaml, key) {
    const lines = yaml.split(/\r?\n/);
    const start = lines.findIndex((l) => l.startsWith(key + ":"));
    if (start < 0) return null;
    const out = [];
    for (const line of lines.slice(start + 1)) {
      if (/^\s+\S/.test(line)) out.push(line.trim());
      else break;
    }
    return out.join(" ") || null;
  }
  fetch("py/rules/itbis_rates.yaml").then((r) => (r.ok ? r.text() : Promise.reject())).then((yaml) => {
    lawNote.es = foldedScalar(yaml, "nota_es");
    lawNote.en = foldedScalar(yaml, "note");
    if (lawNote[lang]) $("#law-note").textContent = lawNote[lang];
  }).catch(() => { /* keep the copy in the HTML */ });

  // ------------------------------------------------------------------ worker
  const worker = new Worker("worker.js", { type: "module" });
  worker.onerror = () => {
    bootStarted = false;
    updateWarm();
    if (screen === "loading") showError("boot");
  };
  worker.onmessage = ({ data }) => {
    if (data.type === "progress") {
      if (data.step !== "check") warmStep = STEPS.indexOf(data.step) + 1;
      updateWarm();
      setStep(data.step);
    } else if (data.type === "ready") {
      ready = true;
      updateWarm();
    } else if (data.type === "boot-error") {
      bootStarted = false;
      updateWarm();
      if (screen === "loading") showError("boot");
    } else if (data.type === "result" && data.id === runId) {
      onResult(data.result, data.report);
    }
  };
  function boot() {
    if (bootStarted) return;
    bootStarted = true;
    worker.postMessage({ type: "boot" });
  }
  // Warm up in the background so results feel instant, unless the user asked to save data.
  const saveData = navigator.connection && navigator.connection.saveData;
  if (!saveData) setTimeout(boot, 400);

  function updateWarm() {
    const node = $("#warm");
    if (ready) node.textContent = t("warmReady");
    else if (saveData && !bootStarted) node.textContent = t("warmSaveData");
    else if (bootStarted) node.textContent = t("warmLoading", { n: Math.max(1, warmStep) });
    else node.textContent = "";
  }

  // ------------------------------------------------------------------ screens
  function show(name, { focus } = {}) {
    screen = name;
    for (const s of ["welcome", "loading", "error", "results"]) $("#screen-" + s).hidden = s !== name;
    window.scrollTo({ top: 0 });
    if (focus) focus.focus({ preventScroll: true });
  }

  function setStep(step) {
    const index = STEPS.indexOf(step);
    document.querySelectorAll("#steps li").forEach((li, i) => {
      li.classList.toggle("is-done", i < index);
      li.classList.toggle("is-active", i === index);
    });
    $("#bar-fill").style.transform = `scaleX(${Math.max(0.04, (index + 0.5) / STEPS.length)})`;
  }

  // ------------------------------------------------------------------ files
  const fileInput = $("#file");
  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
    fileInput.value = "";
  });

  function todayIso() {
    const d = new Date();
    return [d.getFullYear(), String(d.getMonth() + 1).padStart(2, "0"), String(d.getDate()).padStart(2, "0")].join("-");
  }

  async function handleFile(file) {
    state.lastFile = file;
    if (!/\.(xlsx|xlsm|csv|txt)$/i.test(file.name)) {
      showError("type");
      return;
    }
    const id = ++runId;
    $("#loading-file").textContent = file.name;
    setStep(ready ? "check" : (warmStep ? STEPS[warmStep - 1] : "python"));
    clearTimeout(loadingTimer);
    // When Python is already loaded the check takes a moment: don't flash a loading screen.
    if (ready) loadingTimer = setTimeout(() => show("loading"), 300);
    else show("loading");
    boot();
    const buffer = await file.arrayBuffer();
    if (id !== runId) return;
    worker.postMessage({ type: "check", id, name: file.name, buffer, today: todayIso() }, [buffer]);
  }

  function onResult(result, report) {
    clearTimeout(loadingTimer);
    if (!result.ok) {
      showError(result.error, result);
      return;
    }
    if (state.reportUrl) URL.revokeObjectURL(state.reportUrl);
    state.reportUrl = URL.createObjectURL(new Blob([report], { type: XLSX }));
    Object.assign(state, { result, sev: "ALL", query: "", rowFilter: "ALL", shown: PAGE });
    $("#search").value = "";
    show("results");
    renderResults({ animate: true });
    $("#results-title").focus({ preventScroll: true });
    const c = result.counts;
    $("#live").textContent = t("announce", { e: c.ERROR, w: c.ADVERTENCIA, i: c.INFO });
  }

  function showError(kind, data = {}) {
    clearTimeout(loadingTimer);
    const titles = { columns: "errColumnsTitle", read: "errReadTitle", type: "errTypeTitle", empty: "errEmptyTitle", boot: "errBootTitle" };
    const body = $("#error-body");
    body.replaceChildren();
    $("#error-title").textContent = t(titles[kind] || "errReadTitle");
    if (kind === "columns") {
      body.append(
        el("p", {}, t("errColumnsBody"), " ", ...data.missing.flatMap((m, i) => [i ? ", " : "", el("code", { text: m })])),
        el("p", {}, t("errColumnsFound"), " ", ...data.found.flatMap((m, i) => [i ? ", " : "", el("code", { text: m })])),
        el("p", { text: t("errColumnsFix") }),
      );
    } else if (kind === "type") {
      body.append(el("p", { text: t("errTypeBody") }));
    } else if (kind === "empty") {
      body.append(el("p", { text: t("errEmptyBody") }));
    } else if (kind === "boot") {
      body.append(el("p", { text: t("errBootBody") }));
    } else {
      body.append(el("p", { text: t("errReadBody") }));
      if (data.detail) body.append(el("p", {}, t("errDetail"), " ", el("code", { text: data.detail })));
    }
    const retry = $("#error-retry");
    retry.textContent = t(kind === "boot" ? "retry" : "tryAgain");
    retry.dataset.kind = kind;
    show("error", { focus: $("#error-title") });
    $("#live").textContent = $("#error-title").textContent;
  }

  $("#error-retry").addEventListener("click", (e) => {
    if (e.currentTarget.dataset.kind === "boot" && state.lastFile) {
      handleFile(state.lastFile);
    } else {
      show("welcome");
      fileInput.click();
    }
  });
  $("#again").addEventListener("click", () => {
    show("welcome");
    $("#pick").focus({ preventScroll: true });
  });
  $("#try-example").addEventListener("click", async () => {
    const res = await fetch("facturas_ejemplo.csv");
    const blob = await res.blob();
    handleFile(new File([blob], "facturas_ejemplo.csv", { type: "text/csv" }));
  });
  $("#download").addEventListener("click", () => {
    if (!state.reportUrl) return;
    const a = el("a", { href: state.reportUrl, download: state.result.reportName });
    document.body.append(a);
    a.click();
    a.remove();
  });

  // ------------------------------------------------------------------ drag & drop
  const drop = $("#drop");
  let dragDepth = 0;
  const hasFiles = (e) => Array.from(e.dataTransfer?.types || []).includes("Files");
  function setDragging(on) {
    const onWelcome = screen === "welcome";
    drop.classList.toggle("is-over", on && onWelcome);
    document.body.classList.toggle("is-dragging", on && !onWelcome && screen !== "loading");
    $("#drop-title").textContent = t(on && onWelcome ? "dropActive" : $("#drop-title").dataset.i18n);
  }
  window.addEventListener("dragenter", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth += 1;
    setDragging(true);
  });
  window.addEventListener("dragover", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = screen === "loading" ? "none" : "copy";
  });
  window.addEventListener("dragleave", (e) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) setDragging(false);
  });
  window.addEventListener("drop", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (!file || screen === "loading") return;
    drop.classList.add("is-dropped");
    setTimeout(() => drop.classList.remove("is-dropped"), 120);
    handleFile(file);
  });

  // ------------------------------------------------------------------ results
  function rowsText(filas) {
    const shown = filas.slice(0, 8);
    const label = filas.length === 1 ? t("row") : t("rows");
    const parts = [label, " "];
    shown.forEach((f, i) => {
      if (i) parts.push(", ");
      parts.push(el("button", { type: "button", class: "linklike rowlink", "aria-label": t("goRow", { n: f }), onclick: () => jumpToRow(f), text: String(f) }));
    });
    if (filas.length > 8) parts.push(" " + t("andMore", { n: filas.length - 8 }));
    return parts;
  }

  function renderResults({ animate }) {
    const r = state.result;
    $("#results-title").textContent = t("resultsTitle", { file: r.file });
    $("#results-meta").textContent = t("meta", { date: r.today, verified: r.rates.last_verified, age: r.rates.age_days });
    $("#stale").hidden = !r.rates.stale;
    $("#stale").textContent = t("stale", { age: r.rates.age_days });

    const c = r.counts;
    const diff = r.itbis.diff;
    const cards = [
      ["cardInvoices", fmtInt(r.invoices), t("cardInvoicesSub", { lines: fmtInt(r.lines), ok: fmtInt(r.lines_ok) }), "var(--ink)", "var(--ink)"],
      ["cardErrors", fmtInt(c.ERROR), t("cardErrorsSub"), "var(--err)", c.ERROR ? "var(--err)" : "var(--ink)"],
      ["cardWarnings", fmtInt(c.ADVERTENCIA), t("cardWarningsSub"), "var(--warn)", c.ADVERTENCIA ? "var(--warn)" : "var(--ink)"],
      ["cardInfo", fmtInt(c.INFO), t("cardInfoSub"), "var(--info)", "var(--ink)"],
      ["cardItbis", money(diff), t("cardItbisSub", { charged: money(r.itbis.charged), expected: money(r.itbis.expected) }),
        "var(--accent)", Math.abs(diff) > 0.004 ? "var(--err)" : "var(--ok)"],
    ];
    $("#cards").replaceChildren(...cards.map(([label, value, sub, tone, valueColor], i) => el("li", {
      class: "card" + (animate ? " is-entering" : ""),
      style: `--tone:${tone};--value:${valueColor};--i:${i}`,
    }, el("p", { class: "card-label", text: t(label) }),
      el("p", { class: "card-value" + (label === "cardItbis" ? " is-money" : ""), text: value }),
      el("p", { class: "card-sub", text: sub }))));

    const todo = $("#todo");
    todo.replaceChildren(...r.actions.map((a, i) => el("li", { class: `tone-${a.severidad}` },
      el("span", { class: "todo-n", text: String(i + 1).padStart(2, "0") }),
      el("div", {},
        el("span", { class: "todo-title", text: ruleTitle(a.regla, a.titulo) + ". " }),
        el("span", { text: ruleAction(a.regla, a.accion) }),
        el("span", { class: "todo-rows" }, el("span", { class: "sev", text: t("sev")[a.severidad] }), " ", ...rowsText(a.filas))))));
    if (!r.actions.length) todo.replaceChildren(el("li", { class: "todo-none", text: t("todoNone") }));

    renderSeverityChips();
    renderProblems({ animate });
    renderRowChips();
    renderTable();
  }

  function chip(label, n, pressed, tone, onclick) {
    return el("button", { type: "button", class: "chip", "aria-pressed": String(pressed), style: tone ? `--tone:${tone}` : null, onclick },
      tone ? el("span", { class: "swatch", "aria-hidden": "true" }) : null, label, " ", el("span", { class: "n", text: fmtInt(n) }));
  }

  function renderSeverityChips() {
    const c = state.result.counts;
    const total = SEVERITIES.reduce((s, k) => s + c[k], 0);
    const tones = { ERROR: "var(--err)", ADVERTENCIA: "var(--warn)", INFO: "var(--info)" };
    const pick = (sev) => () => {
      state.sev = sev;
      state.shown = PAGE;
      renderSeverityChips();
      renderProblems({ animate: true });
    };
    $("#sev-chips").replaceChildren(
      chip(t("all"), total, state.sev === "ALL", null, pick("ALL")),
      ...SEVERITIES.filter((s) => c[s]).map((s) => chip(t("sevPlural")[s], c[s], state.sev === s, tones[s], pick(s))),
    );
  }

  function filteredIssues() {
    const q = fold(state.query.trim());
    return state.result.issues.filter((i) => {
      if (state.sev !== "ALL" && i.severidad !== state.sev) return false;
      if (!q) return true;
      return [i.fila, i.ncf, i.regla, i.mensaje, i.titulo, RULES_EN[i.regla]?.[0]].some((v) => fold(v).includes(q));
    });
  }

  function renderProblems({ animate }) {
    const list = filteredIssues();
    const total = state.result.issues.length;
    $("#problems-count").textContent = list.length === total ? fmtInt(total) : `${fmtInt(list.length)} / ${fmtInt(total)}`;
    const box = $("#problems");
    const head = el("div", { class: "problem-row problem-head", "aria-hidden": "true" },
      el("span", { text: t("hSeverity") }), el("span", { text: t("hRow") }), el("span", { text: t("hNcf") }),
      el("span", { text: t("hProblem") }), el("span", { class: "h-do", text: t("hDo") }));
    const rows = list.slice(0, state.shown).map((i, n) => {
      return el("div", {
        class: `problem-row tone-${i.severidad}` + (animate && n < 24 ? " is-entering" : ""),
        role: "listitem", style: `--i:${Math.min(n, 24)}`,
      },
      el("div", {}, el("span", { class: "sev", text: t("sev")[i.severidad] })),
      el("div", {}, el("span", { class: "p-cell-label", text: t("hRow") }),
        i.fila ? el("button", { type: "button", class: "linklike rowlink", "aria-label": t("goRow", { n: i.fila }), onclick: () => jumpToRow(i.fila), text: String(i.fila) }) : "—"),
      el("div", { class: "ncf" }, el("span", { class: "p-cell-label", text: t("hNcf") }), i.ncf || "—"),
      el("div", { class: "p-main" },
        el("p", { class: "p-title", text: ruleTitle(i.regla, i.titulo) }),
        lang === "en" ? el("span", { class: "p-msg-label", text: t("detailEs") }) : null,
        el("p", { class: "p-msg", lang: "es", text: i.mensaje })),
      el("div", { class: "p-do" }, el("span", { class: "p-cell-label", text: t("hDo") }), ruleAction(i.regla, i.accion)));
    });
    box.replaceChildren(head, ...rows);
    box.hidden = !list.length;
    $("#problems-empty").hidden = list.length > 0 || total === 0;
    if (total === 0) box.hidden = true;
    const more = $("#problems-more");
    const left = list.length - state.shown;
    more.hidden = left <= 0;
    more.textContent = t("more", { n: fmtInt(Math.min(left, PAGE)) });
  }
  $("#problems-more").addEventListener("click", () => {
    state.shown += PAGE;
    renderProblems({ animate: false });
  });
  let searchTimer = null;
  $("#search").addEventListener("input", (e) => {
    state.query = e.target.value;
    state.shown = PAGE;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => renderProblems({ animate: false }), 80);
  });

  function renderRowChips() {
    const counts = { ERROR: 0, ADVERTENCIA: 0, OK: 0 };
    for (const row of state.result.rows) counts[row.estado] += 1;
    const tones = { ERROR: "var(--err)", ADVERTENCIA: "var(--warn)", OK: "var(--ok)" };
    const pick = (f) => () => {
      state.rowFilter = f;
      renderRowChips();
      renderTable();
    };
    $("#row-chips").replaceChildren(
      chip(t("allRows"), state.result.rows.length, state.rowFilter === "ALL", null, pick("ALL")),
      ...["ERROR", "ADVERTENCIA", "OK"].filter((s) => counts[s]).map((s) => chip(t("sevPlural")[s], counts[s], state.rowFilter === s, tones[s], pick(s))),
    );
  }

  function formatCell(value, numeric) {
    if (value == null || value === "") return null;
    if (typeof value === "number") {
      return Number.isInteger(value) && !numeric ? String(value) : value.toLocaleString("en-US", { maximumFractionDigits: 2, minimumFractionDigits: numeric && !Number.isInteger(value) ? 2 : 0 });
    }
    return String(value);
  }

  function renderTable() {
    const r = state.result;
    const amountCols = new Set(["cantidad", "precio_unitario", "itbis", "total"].map((k) => r.colmap[k]));
    const ncfCols = new Set([r.colmap.ncf, r.colmap.ncf_modificado]);
    const rows = r.rows.filter((row) => state.rowFilter === "ALL" || row.estado === state.rowFilter);
    $("#invoices-count").textContent = rows.length === r.rows.length ? fmtInt(r.rows.length) : `${fmtInt(rows.length)} / ${fmtInt(r.rows.length)}`;
    const thead = el("thead", {}, el("tr", {},
      el("th", { class: "c-fila", scope: "col", text: t("hRow") }),
      ...r.columns.map((c) => el("th", { scope: "col", class: amountCols.has(c) ? "num" : null, text: c })),
      el("th", { scope: "col", text: t("hStatus") })));
    const tbody = el("tbody", {}, ...rows.slice(0, MAX_ROWS).map((row) => el("tr", { class: `st-${row.estado}`, "data-fila": row.fila, tabindex: "-1" },
      el("th", { class: "c-fila", scope: "row", text: String(row.fila) }),
      ...row.values.map((v, i) => {
        const col = r.columns[i];
        const text = formatCell(v, amountCols.has(col));
        return el("td", {
          class: [amountCols.has(col) ? "num" : "", ncfCols.has(col) ? "ncf-col" : "", text == null ? "blank" : ""].join(" ").trim() || null,
          title: text && text.length > 40 ? text : null,
        }, text ?? "—");
      }),
      el("td", {}, el("span", { class: `sev tone-${row.estado}`, text: t("sev")[row.estado] })))));
    $("#invoices").replaceChildren(thead, tbody);
    const note = $("#invoices-note");
    note.hidden = rows.length <= MAX_ROWS;
    note.textContent = t("rowsLimited", { shown: fmtInt(MAX_ROWS), total: fmtInt(rows.length) });
  }

  function jumpToRow(fila) {
    let tr = document.querySelector(`#invoices tr[data-fila="${fila}"]`);
    if (!tr && state.rowFilter !== "ALL") {
      state.rowFilter = "ALL";
      renderRowChips();
      renderTable();
      tr = document.querySelector(`#invoices tr[data-fila="${fila}"]`);
    }
    if (!tr) return;
    const behavior = reducedMotion.matches ? "auto" : "smooth";
    tr.closest(".block").scrollIntoView({ behavior, block: "start" });
    const wrap = tr.closest(".table-wrap");
    wrap.scrollTo({ top: tr.offsetTop - wrap.clientHeight / 2 + tr.offsetHeight, behavior });
    tr.classList.add("is-flash");
    requestAnimationFrame(() => requestAnimationFrame(() => tr.classList.remove("is-flash")));
    tr.focus({ preventScroll: true });
  }

  applyLanguage();
})();
