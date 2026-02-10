//As of 2026-02-10
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config"); // make sure your tab is named Config
  const sectionTitles = sheet.getRange("A1:A3").getValues().flat();
  const sectionContent = sheet.getRange("B1:B3").getValues().flat();
  const payload = {
    a: sectionTitles[0] || "Section A",
    b: sectionTitles[1] || "Section B",
    c: sectionTitles[2] || "Section C",
    d: sectionContent[0] || "Content A \n Content A \n Content A \n Content A",
    e: sectionContent[1] || "Content B \n Content A \n Content A \n Content A",
    f: sectionContent[2] || "Content C \n Content A \n Content A \n Content A",
    updatedAt: new Date().toISOString()
  };

  const json = JSON.stringify(payload);

  // If a callback is provided, return JSONP (bypasses CORS)
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(`${cb}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Fallback (direct JSON)
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
