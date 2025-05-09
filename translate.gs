// Translates a Google Sheet
/** @OnlyCurrentDoc */

function ceviri() {
  var spreadsheet     = SpreadsheetApp.getActive();
  var activeSheet     = spreadsheet.getActiveSheet();
  var actualSheetName = activeSheet.getName();

  var ui = SpreadsheetApp.getUi();

  // Sayfa adı alınır (A1 hücresi)
  var sheetName = activeSheet.getRange('B2').getValue();
  if (!sheetName) {
    ui.alert("Hata: A1 hücresinde geçerli bir sayfa adı bulunamadı. Lütfen bir ad girin.");
    return;
  }

  // Aynı isimde bir sayfa zaten var mı?
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === sheetName) {
      ui.alert("Hata: '" + sheetName + "' isminde bir sayfa zaten mevcut. Lütfen farklı bir ad girin.");
      return;
    }
  }

  // A sütununda veri içeren son satırı bul
  var lastRow = activeSheet
                  .getRange(activeSheet.getMaxRows(), 1)
                  .getNextDataCell(SpreadsheetApp.Direction.UP)
                  .getRow();

  // Aktif sayfayı kopyala
  var newSheet = spreadsheet.duplicateActiveSheet();
  newSheet.setName(sheetName);

  // Kaynak dili al
  var response1 = ui.prompt('Veri Seçimi', 'Lütfen kaynak dili yazın (sadece "tr" veya "en")', ui.ButtonSet.OK_CANCEL);
  var source    = response1.getResponseText().toLowerCase();

  if (response1.getSelectedButton() !== ui.Button.OK || (source !== "tr" && source !== "en")) {
    ui.alert("İşlem iptal edildi: Geçerli bir kaynak dili girilmedi. (Sadece 'tr' veya 'en' kabul edilir)");
    spreadsheet.deleteSheet(newSheet);
    return;
  }

  // Hedef dili al
  var response2 = ui.prompt('Veri Seçimi', 'Lütfen hedef dili yazın (sadece "tr" veya "en")', ui.ButtonSet.OK_CANCEL);
  var target    = response2.getResponseText().toLowerCase();

  if (response2.getSelectedButton() !== ui.Button.OK || (target !== "tr" && target !== "en")) {
    ui.alert("İşlem iptal edildi: Geçerli bir hedef dili girilmedi. (Sadece 'tr' veya 'en' kabul edilir)");
    spreadsheet.deleteSheet(newSheet);
    return;
  }

  // Formül
  var formula = '=IF(ISBLANK(INDIRECT("'+ actualSheetName + '!"&VLOOKUP(COLUMN();Sayfa6!A:B;2;FALSE)&ROW()));"";IF(ISTEXT(INDIRECT("'+ actualSheetName +'!"&VLOOKUP(COLUMN();Sayfa6!A:B;2;FALSE)&ROW()));GOOGLETRANSLATE(INDIRECT("' + actualSheetName +'!"&VLOOKUP(COLUMN();Sayfa6!A:B;2;FALSE)&ROW());"' + source + '";"' + target + '");INDIRECT("'+ actualSheetName +'!"&VLOOKUP(COLUMN();Sayfa6!A:B;2;FALSE)&ROW())))'

  // Aralık tanımları
  var startCol = 1;   // A sütunu
  var endCol   = 13;  // M sütunu
  var startRow = 2;
  var endRow   = lastRow;
  // Formülleri hücrelere uygula
  for (var row = startRow; row <= endRow; row++) {
    for (var col = startCol; col <= endCol; col++) {
      newSheet.getRange(row, col).setFormula(formula);
    }
  }

  // Bilgilendirme metni ekle
  if (response1 = 'tr') {
  newSheet.getRange('B1').setValue('Translated by Google Translate.');
  }
  if (response1 = 'en') {
  newSheet.getRange('B1').setValue('Google Translate tarafından çevrilmiştir.');
  }
  // Başarılı işlem bildirimi
  ui.alert("Sayfa başarıyla oluşturuldu: '" + sheetName + "'");
}
