// Ghicește Maneaua – clasament salvat în Google Sheet
// Lipește tot codul ăsta în Extensions → Apps Script, apoi Deploy → New deployment → Web app

const MAX_SCORE = 50;

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(["Data", "Nume", "Scor"]);
  return sh;
}

// salvează un scor nou (trimis de site la final de joc)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || "").replace(/[<>]/g, "").trim().slice(0, 20);
    const score = Number(d.score);
    if (!name || !Number.isInteger(score) || score < 0 || score > MAX_SCORE) {
      return out_({ ok: false });
    }
    sheet_().appendRow([new Date(), name, score]);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// trimite top 10 către site
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const top = rows
    .filter(r => r[1] !== "")
    .sort((a, b) => b[2] - a[2] || new Date(a[0]) - new Date(b[0]))
    .slice(0, 10)
    .map(r => ({ name: String(r[1]), score: Number(r[2]) }));
  return out_({ top });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
