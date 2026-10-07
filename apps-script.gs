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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
