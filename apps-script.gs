// Pegar en Google Sheet > Extensiones > Apps Script
const HOJA = "Pedidos";

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (d.web) return out({ ok: true }); // honeypot: bots

    const nombre = String(d.nombre || "").trim().slice(0, 80);
    const area = String(d.area || "").trim().slice(0, 80);
    const tel = String(d.telefono || "").trim().slice(0, 20);
    const kilos = Number(d.kilos);
    if (!nombre || !area || !tel || !(kilos >= 0.5 && kilos <= 50)) {
      return out({ ok: false, error: "datos inválidos" });
    }

    const ss = SpreadsheetApp.openById("1HrxjOUBZMj07thyBuVEJm213UuOnvCaRxY-81TjGbx0");
    let sh = ss.getSheetByName(HOJA);
    if (!sh) {
      sh = ss.insertSheet(HOJA);
      sh.appendRow(["Fecha", "Nombre", "Área", "Teléfono", "Kilos", "Entregado", "Pagado"]);
      sh.setFrozenRows(1);
    }
    // evita fórmulas inyectadas (=, +, -, @) en celdas de texto
    const safe = s => /^[=+\-@]/.test(s) ? "'" + s : s;
    sh.appendRow([new Date(), safe(nombre), safe(area), "'" + tel, kilos, false, false]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Lista pública de inscritos (sin teléfono; nombre abreviado). Se usa desde la página con JSONP.
function doGet(e) {
  const cb = (e && e.parameter && e.parameter.callback) || "";
  const ss = SpreadsheetApp.openById("1HrxjOUBZMj07thyBuVEJm213UuOnvCaRxY-81TjGbx0");
  const sh = ss.getSheetByName(HOJA);
  const lista = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
      if (!r[1] || !(Number(r[4]) > 0)) return;
      const partes = String(r[1]).trim().split(/\s+/);
      const nombre = partes.length > 1 ? partes[0] + " " + partes[partes.length - 1][0].toUpperCase() + "." : partes[0];
      lista.push({ n: nombre, a: String(r[2]), k: Number(r[4]) });
    });
  }
  const json = JSON.stringify({ ok: true, lista: lista });
  if (/^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + json + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out({ ok: true, lista: lista });
}
