// Module worker: runs the invoice checker (scripts/check_invoices.py) with Pyodide, off the main thread.
// The file arrives as bytes from the page, is written to Pyodide's in-memory file system,
// and the results and the Excel report go back to the page. Nothing is sent to a server.

const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/";
const PY_FILES = [
  "scripts/check_invoices.py",
  "scripts/web_app.py",
  "rules/itbis_rates.yaml",
  "rules/ncf_types.yaml",
];
// Pure-Python packages that are not part of the Pyodide distribution, pinned from PyPI.
const PYPI = ["openpyxl==3.1.5", "et-xmlfile==2.0.0", "python-stdnum==2.2"];
const HOME = "/home/pyodide/dr-facturacion/";

let booting = null;
let runs = 0;

function progress(step) {
  postMessage({ type: "progress", step });
}

async function boot() {
  progress("python");
  const { loadPyodide } = await import(PYODIDE + "pyodide.mjs");
  const py = await loadPyodide({ indexURL: PYODIDE });

  progress("pandas");
  await py.loadPackage(["pandas", "pyyaml", "micropip"]);

  progress("libs");
  const micropip = py.pyimport("micropip");
  await micropip.install(PYPI);

  progress("rules");
  for (const name of PY_FILES) {
    const res = await fetch(new URL("py/" + name, self.location), { cache: "no-cache" });
    if (!res.ok) throw new Error(`${name}: HTTP ${res.status}`);
    const path = HOME + name;
    py.FS.mkdirTree(path.slice(0, path.lastIndexOf("/")));
    py.FS.writeFile(path, new Uint8Array(await res.arrayBuffer()));
  }
  py.runPython(`import sys; sys.path.insert(0, "${HOME}scripts"); import web_app`);
  postMessage({ type: "ready" });
  return py;
}

function ensureBooted() {
  if (!booting) {
    booting = boot().catch((err) => {
      booting = null; // allow a retry
      throw err;
    });
  }
  return booting;
}

self.onmessage = async ({ data }) => {
  if (data.type === "boot") {
    try {
      await ensureBooted();
    } catch (err) {
      postMessage({ type: "boot-error", detail: String(err && err.message || err) });
    }
    return;
  }
  if (data.type !== "check") return;

  let py;
  try {
    py = await ensureBooted();
  } catch (err) {
    postMessage({ type: "boot-error", id: data.id, detail: String(err && err.message || err) });
    return;
  }
  try {
    progress("check");
    const dir = `/tmp/run${++runs}/`;
    py.FS.mkdirTree(dir);
    const safeName = data.name.replace(/[\\/]/g, "_");
    py.FS.writeFile(dir + safeName, new Uint8Array(data.buffer));
    const webApp = py.globals.get("web_app");
    const result = JSON.parse(webApp.check_file_json(dir + safeName, data.today));
    webApp.destroy();
    let report = null;
    if (result.ok) {
      report = py.FS.readFile(result.report);
      result.reportName = result.report.split("/").pop();
    }
    // Free the in-memory copies; the page keeps only the results and the report bytes.
    for (const name of py.FS.readdir(dir)) {
      if (name !== "." && name !== "..") py.FS.unlink(dir + name);
    }
    py.FS.rmdir(dir);
    postMessage({ type: "result", id: data.id, result, report }, report ? [report.buffer] : []);
  } catch (err) {
    postMessage({ type: "result", id: data.id, result: { ok: false, error: "read", detail: String(err && err.message || err) } });
  }
};
